import openpyxl

# Variables
result = {}

# Definitions
def text_process(text):
  return text + '\n'

def check(people, text):
  if people in result:
    content = result[people]
    content += text_process(text)
    result[people] = content
  else:
    result[people] = text_process(text)

# Open Excel
filename = "data.xlsx"
ktm_wb = openpyxl.load_workbook(filename)
ktm = ktm_wb.worksheets[0]

# Read Excel
for col in ktm.columns:

  # Team
  team_name = str(col[0].value)
  count = int(input(team_name + " 인원 수: "))
  kakao_link = str(col[1].value)
 
  # 
  for num in range(2, 2+count):
    people = str(col[num].value)
    check(people, kakao_link)

# Write Result txt
with open('output.txt', 'w') as f:
  for i in result:
    f.write(i + '\n')
    f.write(result[i])
    f.write('\n')

# Done
print("\n----------Done----------")