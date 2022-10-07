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
	def __init__(self, columnName, rowvalue , defaultUser):
		self.columnName = str(columnName).upper()
		self.rowvalue = rowvalue
		self.defaultUser = str(defaultUser).upper()

	def __str__(self):
		if self.columnName == "ROW_ID" and pandas.isnull(self.rowvalue) :
			self.rowvalue = "'"+str(uuid.uuid4().hex).upper()+"'"
		elif self.columnName == "ACTIVE_FLG" and pandas.isnull(self.rowvalue)  :
			self.rowvalue = "'Y'"
		elif self.columnName == "MODIFICATION_NUM" :
			if pandas.isnull(self.rowvalue) :
				self.rowvalue = "1"
		elif self.columnName == "CREATED" :
			if pandas.isnull(self.rowvalue) :
				self.rowvalue = "SYSDATE"
			else :
				self.rowvalue = "TO_DATE('"+str(self.rowvalue)+"', 'dd-mm-yyyy hh24:mi:ss')"
		elif self.columnName == "CREATED_BY" and pandas.isnull(self.rowvalue) :
			self.rowvalue = "'"+self.defaultUser+"'"
		elif self.columnName == "LAST_UPD" :
			if pandas.isnull(self.rowvalue) :
				self.rowvalue = "SYSDATE"
			else :
				self.rowvalue = "TO_DATE('"+str(self.rowvalue)+"', 'dd-mm-yyyy hh24:mi:ss')"
		elif self.columnName == "LAST_UPD_BY" and pandas.isnull(self.rowvalue) :
			self.rowvalue = "'"+self.defaultUser+"'"
		elif self.columnName == "GROUP_TYPE" and pandas.isnull(self.rowvalue) :
			self.rowvalue = "'CONFIG'"
		else :
			if pandas.isnull(self.rowvalue) :
				self.rowvalue = "''"
			else :
				self.rowvalue = "'"+str(self.rowvalue)+"'"
		return str(self.rowvalue)
	

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
		filenameSql = "{}{}.sql".format(outputdir,fileName+"_"+sheet_name)
		write_file = open(filenameSql, "w")
		for i, _ in data.iterrows():
			field_names = ", ".join(list(data.columns))
			field_names = str(field_names).upper()
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