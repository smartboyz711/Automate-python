import sys
import xml.dom.minidom

print("Scirpt Change Format XML SffRequest From table sff_service_request For Testing PVT/SIT/Production Environment")
print("===========================================================================================")
print("Create BY : Theedanai Poomilamnao 03/10/2022")
print("===========================================================================================")

try :
    filein = input('Input XML File Name (FileName.xml): ')
    f = open(filein,'r')
    newFileData = f.read()
    f.close()

    newFileData = newFileData.replace('<SffRequest>','<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ws="http://th/co/ais/sff/ws" xmlns:java="java:th.co.ais.sff.domain.gm.vo.ws"><soapenv:Header/><soapenv:Body><ws:ExecuteService><ws:sffRequest>')
    newFileData = newFileData.replace('</SffRequest>','</ws:sffRequest></ws:ExecuteService></soapenv:Body></soapenv:Envelope>')
    newFileData = newFileData.replace('<event>','<java:Event>')
    newFileData = newFileData.replace('</event>','</java:Event>')
    newFileData = newFileData.replace('<event/>','<java:Event/>')
    newFileData = newFileData.replace('<parameterList>','<java:ParameterList>')
    newFileData = newFileData.replace('</parameterList>','</java:ParameterList>')
    newFileData = newFileData.replace('<parameterList/>','<java:ParameterList/>')
    newFileData = newFileData.replace('<ParameterList>','<java:ParameterList>')
    newFileData = newFileData.replace('</ParameterList>','</java:ParameterList>')
    newFileData = newFileData.replace('<ParameterList/>','<java:ParameterList/>')
    newFileData = newFileData.replace('<Parameter>','<java:Parameter>')
    newFileData = newFileData.replace('</Parameter>','</java:Parameter>')
    newFileData = newFileData.replace('<Parameter/>','<java:Parameter/>')
    newFileData = newFileData.replace('<name>','<java:Name>')
    newFileData = newFileData.replace('</name>','</java:Name>')
    newFileData = newFileData.replace('<name/>','<java:Name/>')
    newFileData = newFileData.replace('<value>','<java:Value>')
    newFileData = newFileData.replace('</value>','</java:Value>')
    newFileData = newFileData.replace('<value/>','<java:Value/>')
    newFileData = newFileData.replace('<parameterType>','<java:ParameterType>')
    newFileData = newFileData.replace('</parameterType>','</java:ParameterType>')
    newFileData = newFileData.replace('<parameterType/>','<java:ParameterType/>')

    xml = xml.dom.minidom.parseString(newFileData)
    newFileData = xml.toprettyxml()

    fileout = filein.replace('.xml','_SFFXML.xml')
    f = open(fileout,'w')
    f.write(newFileData)
    f.close()
except Exception as e:
    print("An exception occurred Cannot Convert File. : "+str(e))
    print("===========================================================================================")
    input('press any key to Exit program.')
    sys.exit()
    
print('Success Convert File '+filein+' --> '+fileout)
print("===========================================================================================")
input('Press any key to Exit program.')