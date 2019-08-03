import openpyxl

# Start 
print("----------start----------\n")

# Variables
common_text = ''
result = {}
first = True

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

  if first:
    common_text = col[0].value
    first = False
    continue

  # Team
  team_name = str(col[0].value)
  count = int(input(team_name + "\t 인원 수: "))
  team_info = str(col[1].value)
  kakao_link = str(col[2].value)
  text = text_process(team_info) + kakao_link
 
  # Check People
  for num in range(3, 3+count):
    people = str(col[num].value)
    check(people, text)

# Write Result txt
with open('output.txt', 'w', encoding='UTF-8') as f:
  for i in result:
    f.write(text_process(i))
    f.write(text_process(common_text))
    f.write(text_process(result[i]))

# Done
print("\nCheck output.txt file!")
a = input("Press Enter to Exit.")
print("for GT from SH")
print("\n----------Done----------")