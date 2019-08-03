import random

count = int(input('전체 인원 수: '))
winner = int(input('당첨자 수: '))

choices = []
cnt = 0

while True:
  tmp = random.randint(1, count)
  if not tmp in choices:
    choices.append(tmp)
    cnt += 1
  if cnt >= winner:
    break

print(choices)