import pandas as pd

codes = []

print("코드 생성 중...")

# 1단계: 단일 알파벳 + 2자리 숫자 (A01~Z99)
print("1단계: A01~Z99 생성 중...")
for ch in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
    for i in range(1, 100):
        codes.append(f"{ch}{i:02d}")

# 2단계: 두 글자 알파벳 + 1자리 숫자 (AA1~ZZ9)
print("2단계: AA1~ZZ9 생성 중...")
for ch1 in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
    for ch2 in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for i in range(1, 10):
            codes.append(f"{ch1}{ch2}{i}")

# 3단계: 알파벳 + 숫자 + 알파벳 (A1A~Z9Z)
print("3단계: A1A~Z9Z 생성 중...")
for ch1 in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
    for i in range(1, 10):
        for ch2 in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            codes.append(f"{ch1}{i}{ch2}")

# 4단계: 숫자 + 영문 + 숫자 (1A1~9Z9)
print("4단계: 1A1~9Z9 생성 중...")
for num1 in range(1, 10):
    for ch in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        for num2 in range(1, 10):
            codes.append(f"{num1}{ch}{num2}")

# 총 개수 계산
total = len(codes)
# 1단계: 26 * 99 = 2,574개
# 2단계: 26 * 26 * 9 = 6,084개
# 3단계: 26 * 9 * 26 = 6,084개
# 4단계: 9 * 26 * 9 = 2,106개
# 합계: 16,848개
expected = 26 * 99 + 26 * 26 * 9 + 26 * 9 * 26 + 9 * 26 * 9

print(f"\n생성 완료!")
print(f"생성된 코드 개수: {total:,}개")
print(f"예상 개수: {expected:,}개")
print(f"\n첫 10개: {codes[:10]}")
print(f"마지막 10개: {codes[-10:]}")

# Excel 파일 생성
print(f"\nExcel 파일 생성 중...")
df = pd.DataFrame(codes, columns=["Code"])
df.to_excel("codes_A01_to_9Z9.xlsx", index=False)
print(f"완료! codes_A01_to_9Z9.xlsx 파일이 생성되었습니다.")

