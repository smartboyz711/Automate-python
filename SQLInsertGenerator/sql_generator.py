import errno
import os
import sys
import pandas
import uuid

print("========================================================================================")
print("Scirpt generator sql insert from excel file")
print("Create By : theedanai Poomilamnao 07/10/2022")
print("========================================================================================")
print()

def print_exit_program():
	print()
	print("========================================================================================")
	input("press any key to Exit program.")
	sys.exit()

class default_table_value :
	def __init__(self, columnName, rowValue , defaultUser):
		self.columnName = str(columnName).upper()
		self.rowValue = rowValue
		self.defaultUser = str(defaultUser).upper()

	def __str__(self):
		if self.columnName == "ROW_ID" and pandas.isnull(self.rowValue) :
			self.rowValue = "'"+str(uuid.uuid4().hex).upper()+"'"
		elif self.columnName == "MODIFICATION_NUM" :
			if pandas.isnull(self.rowValue) :
				self.rowValue = "1"
		elif self.columnName in "CREATED_BY, LAST_UPD_BY" and pandas.isnull(self.rowValue) :
			self.rowValue = "'"+self.defaultUser+"'"
		elif self.columnName in "CREATED, LAST_UPD, STATUS_DT":
			if pandas.isnull(self.rowValue) :
				self.rowValue = "SYSDATE"
			else :
				self.rowValue = "TO_DATE('"+str(self.rowValue)+"', 'yyyy-mm-dd hh24:mi:ss')"
		elif self.columnName == "GROUP_TYPE" and pandas.isnull(self.rowValue) :
			self.rowValue = "'CONFIG'"
		elif self.columnName == "ACTIVE_FLG" and pandas.isnull(self.rowValue)  :
			self.rowValue = "'Y'"
		else :
			if pandas.isnull(self.rowValue) :
				self.rowValue = "NULL"
			else :
				self.rowValue = "'"+str(self.rowValue)+"'"
		return str(self.rowValue)

try :
	fileIn : str = input("Input excel File Name (FileName.xlsx) : ")
	print()
	defaultUser : str =  input("Input Default User : ")
	print()
	if not fileIn:
		print("Input excel File Name is required, try again next time.")
		print_exit_program()
	if not defaultUser :
		print("Default User is required, try again next time.")
		print_exit_program()

	file = pandas.ExcelFile(fileIn)

	try:
		fileName = fileIn.replace('.xlsx','')
		fileName = fileName.replace('.xls','')
		fileName = fileName.replace('.xlsm','')
		fileName = fileName.replace('.xlsb','')
		fileName = fileName.replace('.odf','')
		fileName = fileName.replace('.ods','')
		fileName = fileName.replace('.odt','')
		os.makedirs(fileName)
		outputdir = "{}/".format(fileName)
	except OSError as exc:
		# If directory is exists use this directory
		if exc.errno == errno.EEXIST:
			outputdir = "{}/".format(fileName)

	for sheet_name in file.sheet_names:
		data = file.parse(sheet_name)
		sheet_name = sheet_name.upper()
		filenameSql = "{}{}.sql".format(outputdir,fileName+"_"+sheet_name)
		write_file = open(filenameSql, "w")
		for i, _ in data.iterrows():
			field_names = ", ".join(list(data.columns))
			field_names = field_names.upper()
			rows = list()
			for column in data.columns:
				defaultRowValue = default_table_value(column,data[column][i],defaultUser)
				rows.append(str(defaultRowValue))
			row_values = ", ".join(rows)
			write_file.write("INSERT INTO {} ({})\nVALUES ({});\n".format(sheet_name, field_names, row_values))
		write_file.write("COMMIT;")
		write_file.close()
		print("Success Convert File Excel "+fileIn+" ===> "+filenameSql)
	print_exit_program()
except Exception as e:
	print("An exception occurred Cannot Generator sql. : "+str(e))
	print_exit_program()