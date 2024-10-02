from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment

dic = {1 : [2],
       3 : [4],
       5 : [6]
       }

dic[3].append(44)

print(dic[3])