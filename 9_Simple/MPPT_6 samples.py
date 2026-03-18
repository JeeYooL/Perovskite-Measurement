#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Keithley 2461 MPPT Control Script
--------------------------------------------------
[실행 방법 - 가상환경 'mppt' 사용 시]

1. 터미널(CMD/PowerShell) 열기
2. 가상환경 활성화:
   - Windows: mppt\\Scripts\\activate
   - Mac/Linux: source mppt/bin/activate
3. 라이브러리 설치 (최초 1회):
   pip install -r requirements.txt
4. 실행:
   python mppt_keithley2461.py
--------------------------------------------------
"""

import pyvisa
import time
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# ==========================================
# 사용자 설정 (USER CONFIGURATION)
# ==========================================
GPIB_ADDRESS = 'GPIB0::18::INSTR'  # 장비의 GPIB 주소 (Menu > Communication에서 확인)
MAX_VOLTAGE = 2.0         # 안전을 위한 최대 인가 전압 (Solar Cell Voc보다 약간 높게 설정)
MAX_CURRENT = 1.0         # 전류 제한 (Compliance) [A]
STEP_SIZE = 0.05          # P&O 알고리즘 전압 변경 스텝 [V]
SAMPLE_INTERVAL = 0.5     # 측정 간격 [초]
DATA_FILENAME = f"mppt_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

# ==========================================
# MPPT 알고리즘 클래스 (Perturb & Observe)
# ==========================================
class MPPTController:
    def __init__(self, step_size):
        self.step_size = step_size
        self.prev_power = 0.0
        self.prev_voltage = 0.0
        self.direction = 1  # 1: 전압 증가, -1: 전압 감소

    def get_next_voltage(self, current_voltage, current_power):
        """
        P&O (Perturb and Observe) 알고리즘 로직
        이전 전력과 현재 전력을 비교하여 다음 전압을 결정함
        """
        delta_p = current_power - self.prev_power
        
        # 전력이 증가했다면 -> 같은 방향으로 계속 이동
        if delta_p > 0:
            pass # direction 유지
        # 전력이 감소했다면 -> 방향 반대로 전환
        else:
            self.direction *= -1

        # 다음 목표 전압 계산
        next_voltage = current_voltage + (self.direction * self.step_size)
        
        # 상태 업데이트
        self.prev_power = current_power
        self.prev_voltage = current_voltage
        
        return next_voltage

# ==========================================
# 메인 프로그램
# ==========================================
def main():
    rm = pyvisa.ResourceManager()
    keithley = None
    data_list = []

    try:
        print(f"Connecting to {GPIB_ADDRESS}...")
        keithley = rm.open_resource(GPIB_ADDRESS)
        print(f"Connected: {keithley.query('*IDN?')}")

        # --- 장비 초기화 및 설정 (SCPI 명령어) ---
        keithley.write('*RST')                  # 장비 리셋
        keithley.write('SOUR:FUNC VOLT')        # Source 모드: Voltage
        keithley.write('SENS:FUNC "CURR"')      # Sense 모드: Current
        keithley.write('SOUR:VOLT:RANG:AUTO ON') # 전압 오토 레인지
        keithley.write('SENS:CURR:RANG:AUTO ON') # 전류 오토 레인지
        keithley.write(f'SOUR:VOLT:ILIM {MAX_CURRENT}') # 전류 제한 설정
        keithley.write('SENS:CURR:NPLC 1')      # 측정 속도 설정 (1 PLC = 보통 속도/정확도)
        
        # 4-wire (Remote Sense) 사용 시 아래 주석 해제
        # keithley.write('SENS:CURR:RSEN ON') 

        # --- 측정 시작 ---
        print("\nStarting MPPT Tracking... (Press Ctrl+C to stop)")
        mppt = MPPTController(STEP_SIZE)
        
        current_v_set = 0.0 # 0V부터 시작
        keithley.write(f'SOUR:VOLT {current_v_set}')
        keithley.write('OUTP ON') # 출력 켜기
        
        start_time = time.time()

        while True:
            # 1. 측정 (Measure)
            # Keithley 2461은 "READ?" 명령 시 기본적으로 "Voltage, Current, ..." 순서로 반환될 수 있음
            # 정확한 포맷팅을 위해 "MEAS:VOLT?;:MEAS:CURR?" 방식을 사용하거나 포맷 지정 필요
            # 여기서는 간단히 READ? 후 파싱합니다.
            
            readings = keithley.query('READ? "defbuffer1", SOUR, READ').split(',')
            # 2461의 READ? 응답은 설정에 따라 다르지만 보통 [Voltage, Current, ...] 순임
            # 안전하게 Source값과 Measure값을 명시적으로 가져옵니다.
            
            meas_v = float(readings[0]) # 측정된 전압
            meas_i = float(readings[1]) # 측정된 전류
            
            # 태양전지 전류는 보통 (-)로 측정되거나 Sink 모드임. 
            # 전력 계산을 위해 절대값 혹은 부호 처리를 합니다.
            # 2461에서 Source Voltage, Measure Current (Sink)일 때 전류는 (-)로 나올 수 있음.
            power = meas_v * abs(meas_i) 

            # 2. 데이터 저장
            elapsed = time.time() - start_time
            print(f"Time: {elapsed:.1f}s | V: {meas_v:.4f}V | I: {meas_i:.4f}A | P: {power:.5f}W")
            
            data_list.append({
                'Time': elapsed,
                'Voltage': meas_v,
                'Current': meas_i,
                'Power': power
            })

            # 3. MPPT 알고리즘 적용 (다음 전압 결정)
            next_v = mppt.get_next_voltage(current_v_set, power)

            # 4. 안전 한계 체크
            if next_v > MAX_VOLTAGE:
                next_v = MAX_VOLTAGE
                mppt.direction = -1 # 강제로 방향 전환
            elif next_v < 0:
                next_v = 0
                mppt.direction = 1

            # 5. 전압 인가
            current_v_set = next_v
            keithley.write(f'SOUR:VOLT {current_v_set}')
            
            time.sleep(SAMPLE_INTERVAL)

    except KeyboardInterrupt:
        print("\nTest stopped by user.")
    except Exception as e:
        print(f"\nError occurred: {e}")
    finally:
        # --- 안전 종료 절차 ---
        if keithley:
            print("Turning off output...")
            keithley.write('OUTP OFF') # 출력 끄기
            keithley.write('ABORT')    # 동작 중지
            keithley.close()
            
        # 데이터 파일 저장
        if data_list:
            df = pd.DataFrame(data_list)
            df.to_csv(DATA_FILENAME, index=False)
            print(f"Data saved to {DATA_FILENAME}")
            
            # 결과 그래프 그리기
            plt.figure(figsize=(10, 6))
            plt.subplot(2, 1, 1)
            plt.plot(df['Time'], df['Power'], 'r-', label='Power (W)')
            plt.ylabel('Power [W]')
            plt.legend()
            plt.grid()
            
            plt.subplot(2, 1, 2)
            plt.plot(df['Time'], df['Voltage'], 'b-', label='Voltage (V)')
            plt.plot(df['Time'], [abs(i) for i in df['Current']], 'g--', label='Current (A)')
            plt.xlabel('Time [s]')
            plt.ylabel('V / I')
            plt.legend()
            plt.grid()
            plt.show()

if __name__ == "__main__":
    main()