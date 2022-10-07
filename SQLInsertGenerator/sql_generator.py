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
	defaultUser = defaultUser.upper()

	try:
		fileName = fileIn.replace('.xlsx','')
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
				columnName = str(column).upper()
				rowvalue = data[column][i]
				if columnName == "ROW_ID" and pandas.isnull(rowvalue) :
					rowvalue = "'"+str(uuid.uuid4().hex).upper()+"'"
				elif columnName == "ACTIVE_FLG" and pandas.isnull(rowvalue)  :
					rowvalue = "'Y'"
				elif columnName == "MODIFICATION_NUM" :
					if pandas.isnull(rowvalue) :
						rowvalue = "1"
				elif columnName == "CREATED" :
					if pandas.isnull(rowvalue) :
						rowvalue = "SYSDATE"
					else :
						rowvalue = "TO_DATE('"+str(rowvalue)+"', 'dd-mm-yyyy hh24:mi:ss')"
				elif columnName == "CREATED_BY" and pandas.isnull(rowvalue) :
					rowvalue = "'"+defaultUser+"'"
				elif columnName == "LAST_UPD" :
					if pandas.isnull(rowvalue) :
						rowvalue = "SYSDATE"
					else :
						rowvalue = "TO_DATE('"+str(rowvalue)+"', 'dd-mm-yyyy hh24:mi:ss')"
				elif columnName == "LAST_UPD_BY" and pandas.isnull(rowvalue) :
					rowvalue = "'"+defaultUser+"'"
				elif columnName == "GROUP_TYPE" and pandas.isnull(rowvalue) :
					rowvalue = "'CONFIG'"
				else :
					if pandas.isnull(rowvalue) :
						rowvalue = "''"
					else :
						rowvalue = "'"+str(rowvalue)+"'"
				rows.append(str(rowvalue))
			row_values = ", ".join(rows)
			write_file.write("INSERT INTO {} ({})\nVALUES ({});\n".format(sheet_name, field_names, row_values))
		write_file.write("COMMIT;")
		write_file.close()
		print("Success Convert File Excel "+fileIn+" ===> "+filenameSql)
	print_exit_program()
except Exception as e:
	print("An exception occurred Cannot Generator sql. : "+str(e))
	print_exit_program()