import errno
import os
import pandas
import uuid

print("========================================================================================")
print("Scirpt generator sql insert from excel file")
print("Create By : theedanai Poomilamnao 07/10/2022")
print("========================================================================================")
print()

def print_line():
	print()
	print("========================================================================================")
	print()

class default_table_value :
	def __init__(self, columnName, rowValue , defaultUser):
		self.columnName = str(columnName).upper()
		self.rowValue = rowValue
		self.defaultUser = str(defaultUser).upper()

	def __str__(self):
		if self.columnName == "ROW_ID" and pandas.isnull(self.rowValue) :
			self.rowValue = "'"+str(uuid.uuid4().hex).upper()+"'"
		elif self.columnName in ["MODIFICATION_NUM","ORDER_BY"] :
			if pandas.isnull(self.rowValue) :
				self.rowValue = "0"
		elif self.columnName in ["CREATED_BY", "LAST_UPD_BY"] and pandas.isnull(self.rowValue) :
			self.rowValue = "'"+self.defaultUser+"'"
		elif  self.columnName in ["CREATED", "LAST_UPD", "STATUS_DT","RULE_START_DT","EFF_START_DT","CRITERIA_START_DT"] :
			if pandas.isnull(self.rowValue) :
				self.rowValue = "SYSDATE"
			else :
				self.rowValue = f"TO_DATE('{str(self.rowValue)}', 'yyyy-mm-dd hh24:mi:ss')"
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

def main() :
	while True :
		try :
			fileIn : str = input("Input excel File Name (FileName.xlsx) : ")
			print()
			defaultUser : str =  input("Input Default User : ")
			print()
			if not fileIn:
				print("Input excel File Name is required, try again next time.")
				print_line()
				continue
			if not defaultUser :
				print("Default User is required, try again next time.")
				print_line()
				continue

			file = pandas.ExcelFile(fileIn)

			try:
				fileName = fileIn.replace('.xlsx','')
				fileName = fileName.replace('.xls','')
				os.makedirs(fileName)
				outputdir = "{}/".format(fileName)
			except OSError as exc:
				# If directory is exists use this directory
				if exc.errno == errno.EEXIST:
					outputdir = "{}/".format(fileName)
     
			filenameSql = "{}{}.sql".format(outputdir,fileName)
			write_file = open(filenameSql, "w", encoding="utf-8")
			for sheet_name in file.sheet_names:
				data = file.parse(sheet_name)
				sheet_name = sheet_name.upper()
				write_file.write(f"/*--------------------{sheet_name}--------------------*/\n")
				for i, _ in data.iterrows():
					field_names = ", ".join(list(data.columns))
					field_names = field_names.upper()
					rows = []
					for column in data.columns:
						defaultRowValue = default_table_value(column,data[column][i],defaultUser)
						rows.append(str(defaultRowValue))
					row_values = ", ".join(rows)
					write_file.write("INSERT INTO {} ({})\nVALUES ({});\n".format(sheet_name, field_names, row_values))
				write_file.write("COMMIT;\n")
			write_file.close()
			print("Success Convert File Excel "+fileIn+" ===> "+filenameSql)
			print_line()
			continue
		except Exception as e:
			print("An exception occurred Cannot Generator sql. : "+str(e))
			print_line()
			continue

if __name__ == "__main__":
    main()