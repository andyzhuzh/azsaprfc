import json
import os
import re
import pyrfc
import sys
import pandas
from dataclasses import dataclass,field,asdict

@dataclass(frozen=True)
class ServerParameters:
    hostname:str
    client:str 
    user:str 
    password:str 
    sysnr:str
    language:str 
    router:str 
    
          
class SapServer:
    '''SAP RFC 连接服务器的相关基类'''
    # 连接参数样例
    CONNECT_PARAMETERS = {
        'ASHOST': '',  # 'XXX.XXX.XX.XX'
        'CLIENT':  '',  # '800'
        'SYSNR':  '',  # '00'
        'USER':   '',
        'PASSWD':  '',
        'LANG': 'ZH',
        'SAPROUTER': '',  # '/H/XXX.XX.XX.XX/H/XXX.XX.XXX.XXX/H/'
    }

    def __init__(self, hostname, client, user='', password='', sysnr='00', language='ZH', router='') -> None:
        self.conParameters = {
            'ASHOST': hostname,   
            'CLIENT':  client,  # '800'
            'SYSNR':  sysnr,  # '00'
            'USER':   user,
            'PASSWD':  password,
            'LANG': language,
            'SAPROUTER': router,   
        }
        self.sapObject = None

    def connect(self) -> bool:
        '''调用RFC前需要调用此方法'''
        # 连接SAP
        self.sapObject = pyrfc.Connection(**self.conParameters)
        if self.sapObject:
            if self.sapObject.alive:
                return True
            else:
                return False
        else:
            return False

    def disconnect(self):
        '''调用RFC后调用此方法断开连接'''
        if self.sapObject:
            if self.sapObject.alive:
                self.sapObject.close

    def callRFC(self, functionName, **parameters):
        """ 调用RFC函数
        参数说明：
        :functionName: RFC函数名称，如：ZFM_RTR_GET_MESSAGE
        :parameters: 字典形式
        调用样例:
        parameters = {'IM_MSG' : sqlStatement,
                      'IM_SOURCE': sourcelist } 
        functionName =  'ZFM_RTR_GET_MESSAGE' 
        resultList = [] 
        try:
            self.connect() 
            returnResult  = self.callRFC(functionName,**parameters)  
            self.disconnect() 
            resultBinaryString = returnResult['EV_RESULT']  
            resultRows = returnResult['EV_ROWS'] 
            returnMesage = returnResult['EV_MESSAGE'] 
        except Exception as e: 
            str(e) 
        finally: 
            self.disconnect() 
        """
        try:
            result = self.sapObject.call(functionName, **parameters)
            return result
        except Exception as e:
            raise e

    # call function:RFC_READ_TABLE
    def readTable(self, tableName, conds: list = [], fields: list = [], DELIMITER='|'):
        '''调用READ_TABLE函数获取表内容

        如：readTable('SPFLI',[{"TEXT": "COUNTRYFR EQ 'US' "} ],['CARRID','CITYFROM']) 
        '''
        parameters = {'QUERY_TABLE': tableName,
                      'DELIMITER': DELIMITER,
                      'OPTIONS': conds,
                      'FIELDS': fields,
                      }
        functionName = 'RFC_READ_TABLE'
        try:
            self.connect()
            returnResult = self.callRFC(functionName, **parameters)
            self.disconnect()
            fieldsList = returnResult['FIELDS']
            fieldsnameList = []
            resultData = []
            for field in fieldsList:
                fieldsnameList.append(field.get('FIELDNAME'))
            for dat in returnResult['DATA']:
                waList = dat['WA'].split(DELIMITER)
                wa1 = [wa.strip() for wa in waList]  # 删除文本空格
                rowDict = dict(zip(fieldsnameList, wa1))
                resultData.append(rowDict)
        except Exception as e:
            resultData.append({'info': str(e)})
        finally:
            self.disconnect()
            return resultData

    # import pandas
    # sap.readTable('excelfile.xls','SPFLI',[{"TEXT": "COUNTRYFR EQ 'US' "} ],['CARRID','CITYFROM'])
    def readTableToEXCEL(self,filename,tableName,conds:list = [],fields:list=[],DELIMITER = '|'): 
        fullfilename = filename +'.xlsx'
        parameters = {'QUERY_TABLE':tableName, 
                'DELIMITER':DELIMITER, 
                'OPTIONS':conds, 
                'FIELDS':fields, 
                } 
        functionName =  'RFC_READ_TABLE' 
        returnResult  = self.callRFC(functionName,**parameters) 
        result = pandas.DataFrame([[x.strip() for x in returnResult['DATA'][n]['WA'].split(DELIMITER)] for n in range(len(returnResult['DATA']))], 
                columns=[x['FIELDNAME'] for x in returnResult['FIELDS']] )
        result.to_excel(fullfilename,index=False)
        
    def restExecuteSql(self, sqlStatement):
        '''通过ADT SQL 执行SQL 查询语句

        sqlStatement:SQL 语句
        有ADT 开发权限
        call function:SADT_REST_RFC_ENDPOINT
        '''
        import xml.etree.ElementTree as ET
        request_line = {'METHOD': 'POST',
                        'URI': '/sap/bc/adt/datapreview/freestyle?rowNumber=100&dataAging=true',
                        'VERSION': 'HTTP/1.1',
                        }
        headers = [{'NAME': 'Content-Type',
                    'VALUE': 'text/plain',
                    },
                   {'NAME': 'Accept',
                    'VALUE': 'application/xml, application/vnd.sap.adt.datapreview.table.v1+xml',
                    },
                   {'NAME': 'User-Agent',
                    'VALUE': 'Eclipse/4.28.0.v20230605-0440 (win32; x86_64; Java 17.0.7) ADT/3.36.0 (devedition)',
                    },
                   {'NAME': 'X-sap-adt-profiling',
                    'VALUE': 'server-time',
                    },
                   ]
        request = {'REQUEST_LINE': request_line,
                   'HEADER_FIELDS': headers,
                   'MESSAGE_BODY': bytes(sqlStatement, 'utf-8'),
                   }
        parameters = {'REQUEST': request}
        functionName = 'SADT_REST_RFC_ENDPOINT'
        try:
            self.connect()
            returnResult = self.callRFC(functionName, **parameters)
            self.disconnect()
            response = returnResult['RESPONSE']
            body = response['MESSAGE_BODY']
            bodyString = str(body, encoding='utf8')
            # xlm to list
            content = ET.fromstring(bodyString)
            root = content
            NS = {"dataPreview": r"http://www.sap.com/adt/dataPreview"}
            resultList = []
            resultRow = {}
            columnData = []
            rowCount = 0
            columCount = 0
            # print(root.attrib['xmlns:dataPreview']) #.get('dataPreview'))
            for column in root.findall('dataPreview:columns', namespaces=NS):
                columCount += 1
                col = {'name': None,
                       'data': []}
                for field in column.findall('dataPreview:metadata', namespaces=NS):
                    col['name'] = field.attrib.get(
                        '{'+NS['dataPreview']+'}description')

                    # print(field.attrib.get('{'+NS['dataPreview']+'}description'))
                for dataset in column.findall('dataPreview:dataSet', namespaces=NS):
                    for node in dataset.iter():
                        col['data'].append(node.text)
                        if columCount == 1:
                            rowCount += 1
                columnData.append(col)

            for i in range(rowCount):
                resultRow = {}
                for dat in columnData:
                    resultRow[dat.get('name')] = dat.get('data')[i]
                resultList.append(resultRow)
        except Exception as e:
            resultList.append({'info': str(e)})
        finally:
            self.disconnect()
            return resultList

 
    def restReadSourceBatch(self,filePath,objList:list):
        """_summary_

        Args:
            filePath (string): 路径
            objList (list): 对象列表  'objType' : CLASS,PROG,FUNC,TABL
            objList = [ {'objType':'CLASS',  'objName':'ZTEST'},
                         {'objType':'PROG', 'objName':'YTEST001'},
            ]
        """
        resultList = []
        # contList = [
        #     # {'objType':'CLASS',  #CLASS,PROG,FUNC,TABL
        #     # 'objName':'ZTEST',
        #     # },
        #     {'objType':'PROG',
        #     'objName':'YTEST001',
        #     },
        #     ]
        # paraServer = ConnParams.S4D
        # sapServer = SapServer(**asdict(paraServer))
        if objList == None:
            print('objList is None')
            return None
        for cnt in objList:
            resultDict={}
            rest = self.restReadSource(cnt.get('objType'),cnt.get('objName'))
            fullFile = os.path.join(filePath,cnt.get('objName')+'.txt')
            with open(fullFile,'w',encoding='UTF-8') as f:
                f.write(rest)
            resultDict = {
                'objType' : cnt.get('objType'),
                'objName': cnt.get('objName'),
                'fileName':fullFile,
                          }
            resultList.append(resultDict)
            print( f'{ fullFile },Done!')
        return resultList
        
        
    def restReadSource(self, objType,objName):
        ulr = None
        urltxtsymbol=None
        urltxtselection=None
        retString = None
        if objName == '':
            return None
        if objType == 'CLASS':
            ulr = f'/sap/bc/adt/oo/classes/{objName.lower().strip()}/source/main'
            urltxtsymbol =  f'/sap/bc/adt/textelements/classes/{objName.lower().strip()}/source/symbols'

        elif objType == "PROG":
            ulr = f'/sap/bc/adt/programs/programs/{objName.lower().strip()}/source/main'
            urltxtsymbol =  f'/sap/bc/adt/textelements/programs/{objName.lower().strip()}/source/symbols'
            urltxtselection = f'/sap/bc/adt/textelements/programs/{objName.lower().strip()}/source/selections'

        elif objType == 'FUNC':
            cond = [{"TEXT": f"FUNCNAME EQ '{objName.upper()}'"} ]
            funcs = self.readTable('TFDIR',cond,['PNAME'])
            funcgrp = funcs[0].get('PNAME')[4:]
            ulr = f'/sap/bc/adt/functions/groups/{funcgrp.lower().strip()}/fmodules/{objName.lower().strip()}/source/main'
            urltxtsymbol =  f'/sap/bc/adt/textelements/functiongroups/{funcgrp.lower().strip()}/source/symbols'
            urltxtselection =  f'/sap/bc/adt/textelements/functiongroups/{funcgrp.lower().strip()}/source/selections'
        elif objType == 'TABL':
            retString = self.restReadStructure(objName)
        elif objType == 'INFO':
            retString = self.restReadObjectList(objName)
            
        if ulr != None:
            retString = self.restCallURL(ulr)
            
        retString += '\n\n"=====以下是相关文本信息=====\n\n'
        if urltxtsymbol !=None:
             # text symbols
            retTxtSymbol = self.restCallURL(urltxtsymbol,Accepttype='application/vnd.sap.adt.textelements.symbols.v1')
            txtList=retTxtSymbol.split('\r')
            retString += '\n\n"=====以下是文本元素=====\n'  
            for txt in txtList:
                if txt and txt[:1] != '@':
                    retString += '\n' +  txt
        if urltxtselection !=None:
             # text symbols
            retTxtSymbol = self.restCallURL(urltxtselection,Accepttype='application/vnd.sap.adt.textelements.selections.v1')
            txtList=retTxtSymbol.split('\r')
            retString += '\n\n"=====以下是Selection文本=====\n' 
            for txt in txtList:
                if txt :
                    retString += '\n' +  txt
        return retString
         
        
#     property
# Accept    : application/vnd.sap.adt.repository.objproperties.result.v1+xml
# GET /sap/bc/adt/repository/informationsystem/objectproperties/values?
# uri=/sap/bc/adt/functions/groups/zfg_rtr_f0007/fmodules/zfm_rtr_int_017_process&facet=package&facet=appl
# /sap/bc/adt/programs/programs/zrtrrf0140&facet=package&facet=appl
# /sap/bc/adt/oo/classes/zcl_abap_utilities&facet=package&facet=appl
# <?xml version="1.0" encoding="UTF-8"?>
# +<opr:object expandable="true" type="CLAS/OC" package="ZRTR" name="ZCL_ABAP_UTILITIES" text="工具">

# structure
# /sap/bc/adt/ddic/elementinfo?path=zscud_info
# Accept    : application/vnd.sap.adt.elementinfo+xml
# adtcore:name="zrtrsf00100" adtcore:type="TABL/DS" adtcore:uri="/sap/bc/adt/ddic/structures/zrtrsf00100">

# abap URL list 代码
# DATA(lr_inst) = cl_adt_tools_core_factory=>get_instance( ).
# DATA(lr_urlmapper) = lr_inst->get_uri_mapper( ).
# " CL_ADT_URI_TEMPLATES_SHM_ROOT->BUILD_TEMPLATES
# DATA(type_provider) = cl_wb_registry=>get_objtype_provider( ).
# type_provider->get_objtypes( IMPORTING p_objtype_data = DATA(wbobjtype_data) ).
# LOOP AT wbobjtype_data->mt_wbobjtype ASSIGNING FIELD-SYMBOL(<type>) WHERE pgmid              = 'R3TR'
#                                                                     AND   is_main_subtype_wb = abap_true.
#   DATA ls_type TYPE wbobjtype.
#   ls_type-objtype_tr = <type>-objecttype.
#   ls_type-subtype_wb = <type>-subtype_wb.
# *<type>-description
#   DATA(lv_type) = |{ <type>-objecttype }{ <type>-subtype_wb }|.
#   DATA(uri) = cl_adt_tools_core_factory=>get_instance( )->get_uri_mapper( )->get_adt_object_ref_uri(
#                   name = 'ZCL_RTR_COMMON'
#                   type = ls_type  ).
# ENDLOOP.
    
    def restCallURL(self, url,Accepttype='text/plain'):
        '''通过ADT 获取代码
        有ADT 开发权限
        call function:SADT_REST_RFC_ENDPOINT
        url:
        class source:/sap/bc/adt/oo/classes/{classname.lower().strip()}/source/main  
        report source: /sap/bc/adt/programs/programs/{programname.lower().strip()}/source/main source
        report text: /sap/bc/adt/textelements/programs/{programname.lower().strip()/source/selections?version=workingArea  
                     /sap/bc/adt/textelements/programs/{programname.lower().strip()/source/symbols?version=workingArea 
        funct source: /sap/bc/adt/functions/groups/{funcgrp.lower().strip()}/fmodules/{objName.lower().strip()}/source/main
        /sap/bc/adt/ddic/tables/mara/source/main
        '''
        request_line = {'METHOD': 'GET',
                        'URI': url,
                        'VERSION': 'HTTP/1.1',
                        }
        headers = [ 
                    
                   {'NAME': 'Cache-Control',
                    'VALUE': 'no-cache',
                    },
                   {'NAME': 'Accept',
                    'VALUE': Accepttype,  
                    },
                   {'NAME': 'User-Agent',
                    'VALUE': 'Eclipse/4.28.0.v20230605-0440 (win32; x86_64; Java 17.0.7) ADT/3.38.2 (devedition)',
                    },
                   {'NAME': 'X-sap-adt-profiling',
                    'VALUE': 'server-time',
                    },
                   ]
        request = {'REQUEST_LINE': request_line,
                   'HEADER_FIELDS': headers,
                #    'MESSAGE_BODY': bytes(sqlStatement, 'utf-8'),
                   }
        parameters = {'REQUEST': request}
        functionName = 'SADT_REST_RFC_ENDPOINT'
        try:
            self.connect()
            returnResult = self.callRFC(functionName, **parameters)
            self.disconnect()
            response = returnResult['RESPONSE']
            body = response['MESSAGE_BODY']
            # print(body)
            bodyString1 = str(body, encoding='utf8')
            # bodyString = str(body, encoding='utf8')
            bodyString = str(body, encoding='utf8').replace('\n','')
            # print(bodyString1)
        except Exception as e:
            pass
        finally:
            self.disconnect() 
            return bodyString        

    def restReadStructure(self,objName):
        url = f'/sap/bc/adt/ddic/elementinfo?path={objName.lower().strip()}'
        resultXmlString=self.restCallURL(url ,Accepttype='application/vnd.sap.adt.elementinfo+xml')
        lines = 0
        import xml.etree.ElementTree as ET
        xmlroot = ET.fromstring(resultXmlString)
        NS = {'abapsource':r'http://www.sap.com/adt/abapsource',
                'adtcore':r'http://www.sap.com/adt/core'}
        tabdesc = xmlroot.find("abapsource:documentation",namespaces= NS).text
        fieldstring = f'Talble/Structure:{objName}, Description:{tabdesc}\n'
        # tabdesc = xmlroot.find('{'+NS['abapsource']+'}documentation').text
        resultList=[]
        columns = xmlroot.findall('abapsource:elementInfo' , namespaces= NS )
        for column in columns:
            lines +=1
            if column.attrib.get('{'+NS['adtcore']+'}type') != "TABL/DTF":
                continue
            field={}
            field['fieldName'] = column.attrib.get('{'+NS['adtcore']+'}name')
            field['fieldDesc'] = column.find('abapsource:documentation' , namespaces= NS).text
            # field['fieldDesc'] = column.findall('abapsource:documentation' , namespaces= NS)[0].text
            fieldinfos = column.findall('abapsource:properties' , namespaces= NS)
            for info in fieldinfos[0].findall('abapsource:entry' , namespaces= NS):
                if info.attrib.get('{'+NS['abapsource']+'}key') == 'ddicDataElement':
                    field['dtlElement'] = info.text
                elif info.attrib.get('{'+NS['abapsource']+'}key') == 'ddicDataType':
                    field['dataType'] = info.text
                elif info.attrib.get('{'+NS['abapsource']+'}key') == 'ddicLength':
                    field['dataLen'] = info.text
                elif info.attrib.get('{'+NS['abapsource']+'}key') == 'ddicDecimals':
                    field['dataDecimals'] = info.text
                elif info.attrib.get('{'+NS['abapsource']+'}key') == 'ddicIsKey':
                    field['isKey'] = info.text
                elif info.attrib.get('{'+NS['abapsource']+'}key') == 'parentName':
                    field['parentName'] = info.text                
            if lines == 1:
                fieldstring +=  '\n' + '\t'.join(field.keys())
                fieldstring += '\n=====================================================\n'
            fieldstring +=  '\n' + '\t'.join(field.values())
            resultList.append(field)
        return fieldstring
 
    def restReadObjectList(self,objName):
        url = f'/sap/bc/adt/repository/informationsystem/search?operation=quickSearch&query={objName.lower().strip()}*&maxResults=51'
        resultXmlString=self.restCallURL(url ,Accepttype='application/xml')
        lines = 0
        import xml.etree.ElementTree as ET
        xmlroot = ET.fromstring(resultXmlString)
        NS = { 'adtcore':r'http://www.sap.com/adt/core'}
        # tabdesc = xmlroot.find('{'+NS['abapsource']+'}documentation').text
        fieldstring = f'Searching:{objName}:\n'
        resultList=[]
        objlists = xmlroot.findall('adtcore:objectReference' , namespaces= NS )
        for obj in objlists:
            lines +=1
            result={}
            result['name'] = obj.attrib.get('{'+NS['adtcore']+'}name') 
            result['type'] = obj.attrib.get('{'+NS['adtcore']+'}type') 
            result['description'] = obj.attrib.get('{'+NS['adtcore']+'}description') 
            result['uri'] = obj.attrib.get('{'+NS['adtcore']+'}uri') 
            result['packageName'] = obj.attrib.get('{'+NS['adtcore']+'}packageName') 
            if lines == 1:
                fieldstring +=  '\n' + '\t'.join(result.keys())
                fieldstring += '\n=====================================================\n'
            fieldstring +=  '\n' + '\t'.join(result.values())
            resultList.append(result)
        return fieldstring
        # return resultList
    
        
    def executeProgramDynamic(self, codeFileName,funcID, para1, para2='', para3='', para4=''):
        '''通过函数 ECATT_REMOTE_MODULE_EXECUTE 动态执行各种语句

        需要权限对象S_DEVELOP：
        OBJTYPE:'SCAT', 'ACTVT' '16'.
        或 'OBJTYPE' , 'ECSC'  'ACTVT' '16'.
        执行代码在文件：sapdynamicprogram.abap
        '''

        abapsource = ''
        resultJson = ''
        result = {}
        resultList = []
        current_file_path = os.path.abspath(__file__)  # 当前文件绝对路径
        current_path = os.path.dirname(current_file_path)  # 文件所在的路径
        filenamewithpath = os.path.join(current_path, 'sapdynamicprogram.abap')
        sourcelist = []
        inputPara = []
        filenamewithpath = codeFileName
        with open(filenamewithpath, 'r', encoding='UTF-8') as f:
            abapsource = f.readline()
            source = {'LINE': abapsource.replace("\n", "")}
            sourcelist.append(source)
            while abapsource:
                abapsource = f.readline()
                source = {}
                source = {'LINE': abapsource.replace("\n", "")}
                sourcelist.append(source)
        inputParaDict = {'FUNCTIONID': funcID,
                         'PARA1': para1,
                         'PARA2': para2,
                         'PARA3': para3,
                         'PARA4': para4,
                         }
        inputParaJson = json.dumps(inputParaDict, ensure_ascii=False)
        # 30长度拆分字符到List
        inputParaList = re.findall('.{'+str(30)+'}', inputParaJson)
        inputParaList.append(inputParaJson[(len(inputParaList)*30):])

        for para in inputParaList:
            paradict = {'FLAG': '', 'VALUE': para}
            inputPara.append(paradict)
        parameters = {'PROGRAM': sourcelist,
                      'ECATT_TRANSFER_DATA_CONTAINER': inputPara}
        functionName = 'ECATT_REMOTE_MODULE_EXECUTE'
        resultList = []
        try:
            self.connect()
            returnResult = self.callRFC(functionName, **parameters)
            self.disconnect()
            returnMsg = returnResult.get('FUN_MESSAGE').get('MSGTEXT')
            if returnResult.get('SY_SUBRC') == 0:
                returnCode = '200'
            else:
                returnCode = '201'
            resultTable = returnResult.get('ECATT_TRANSFER_DATA_CONTAINER')
            for res in resultTable:
                resultJson += res.get('VALUE')
            print(resultJson)
            if resultJson:
                resultList = json.loads(resultJson)
            result = {
                'ResultList': resultList,
                'Rows': 0,
                'Code': returnCode,
                'Message': returnMsg,
            }
        except Exception as e:
            # print(str(e))
            result = {
                'ResultList': [],
                'Rows': 0,
                'Code': '201',
                'Message': str(e),
            }
        finally:
            self.disconnect()
            return result

    
    def getFunctionParameters(self,imFunctionName):
        resultList=[]
        try:
            self.connect() 
            func_desc = self.sapObject.get_function_description(imFunctionName)
            for parameter in func_desc.parameters:
                para = {
                "direction":parameter["direction"],
                "name":parameter["name"],
                "parameter_text":parameter["parameter_text"],
                "parameter_type":parameter["parameter_type"],
                "nuc_length":parameter["nuc_length"],
                "uc_length":parameter["uc_length"],
                "decimals":parameter["decimals"],
                "default_value":parameter["default_value"],
                "optional":parameter["optional"],
                "type_name":None ,
                }
                if parameter["type_description"] != None:
                    para['type_name'] = parameter["type_description"].name               
                resultList.append(para)
                # print(para)
        except:
            pass
        finally:
            self.disconnect() 
            return resultList
        
    def getFunctionParameterType(self,imFunctionName,imParameter):
        resultDict={}
        resultFieldList = []
        try:
            self.connect() 
            func_desc = self.sapObject.get_function_description(imFunctionName)
            for parameter in func_desc.parameters:
                if parameter.get("name") == imParameter and parameter["type_description"] != None:
                    
                    for itm in parameter["type_description"].fields:
                        # print(itm)
                        fields ={
                            # 'FieldType':itm['type']
                            'name':itm['name'],
                            'field_type':itm['field_type'],
                            'nuc_offset':itm['nuc_offset'],
                            'nuc_length':itm['nuc_length'],
                            'uc_length':itm['uc_length'],
                            'uc_offset':itm['uc_offset'],
                            'decimals':itm['decimals'],
                            'type_name':itm['type_description'].name if itm['type_description'] !=None else "",
                            }
                        resultFieldList.append(fields)
                    resultDict={
                        'para_name':imParameter,
                        'type_name':parameter["type_description"].name,
                        'fields':resultFieldList
                    }
        except:
            pass
        finally:
            self.disconnect() 
            return resultDict


    def getTableStructure(self,imTableName):
        resultList=[]
        if imTableName == None or imTableName == "" :
            return None
        try:
            self.connect() 
            parameter={
                "TABNAME":   imTableName,
		        "ALL_TYPES": "X",
            }
            functionName = "DDIF_FIELDINFO_GET"
            tabstru = self.callRFC(functionName,**parameter)
            if tabstru.get("DDOBJTYPE") == "TTYP":
                print(tabstru.get("DFIES_WA").get("ROLLNAME"))
                resultList = self.getTableStructure(tabstru.get("DFIES_WA").get("ROLLNAME"))
                return resultList
            for field in  tabstru.get("DFIES_TAB"):
                fields ={
                    'FieldName':  field["FIELDNAME"],
                    'Position': field["POSITION"],
                    'DataType': field["DATATYPE"],
                    'FieldText': field["FIELDTEXT"],
                    'OffSet':  field["OFFSET"],
                    'RollName': field["ROLLNAME"],
                    'Leng':field["LENG"],
                    'Decimals':field["DECIMALS"],
                    'FieldTitle':  field["SCRTEXT_L"],
                    }
                resultList.append(fields)
            
        except:
            pass
        finally:
            self.disconnect() 
            return resultList
    
    def printTableStructure(self,imTableName):
        resultList = self.getTableStructure(imTableName)
        for result in resultList:
            print(result)
    
    def printFunctionParameterType(self,imFunctionName,imParameter):
        resultList = self.getFunctionParameterType(imFunctionName,imParameter)
        print(resultList['type_name'])
        for result in resultList['fields']:
            print(result)    
   
    def printFunctionParameters(self,imFunctionName):
        resultList = self.getFunctionJsonParameters(imFunctionName)
        for result in resultList:
            print(result)             
        
    def getFunctionJsonParameters(self,imFunctionName):
        resultDict = {}
        funcParams = self.getFunctionJsonParameters(imFunctionName)
        importPara = {}
        # importPara = {fp['name']:fp.get('default_value',None)  for fp in funcParams if fp['direction'] == 'RFC_IMPORT'}
        for fp in funcParams:
            if fp['direction'] == 'RFC_IMPORT':
                if fp['type_description'] != None:
                    fieldsdict = {field['name']:None for field in fp['type_description']['structure']}
                    importPara.update({fp['name']:fieldsdict })  
                else:
                    importPara.update({fp['name']:fp.get('default_value',None)})  
        tablePara  = {fp['name']:[{field['name']:None for field in fp['type_description']['structure']}] for fp in funcParams if fp['direction'] == 'RFC_TABLES'}
        for dt in [importPara, tablePara]:
            resultDict.update(dt)
        return  json.dumps(resultDict)
   
class SAPUILandscape:
    def __init__(self,file_name  ) :
        import xml.etree.ElementTree as ET
        self.content = ET.parse(file_name)
        self.Items = []
        self.Services = []
        self.Routers = []
        self.Messageservers= []
        self.ServerList = []
      
        import xml.etree.ElementTree as ET
        self.content = ET.parse(file_name)
        root = self.content.getroot()
        # root[0].tag 
        # root[0].text 
        for item in root.iter('Item'): # root.findall('Item'):
            itm= {
                'uuid':item.get('uuid'),
                'serviceid':item.get('uuid')
            }
            
        
        for item in root.iter('Item'): # root.findall('Item'):
            itm= {
                'uuid':item.get('uuid'),
                'serviceid':item.get('uuid')
            }
            self.Items.append(itm)
        
        for Route in root.iter('Router'): # root.findall('Router'):
            dict = {
                'uuid':Route.get('uuid'),
                'name':Route.get('name'),
                'description':Route.get('description'),
                'router':Route.get('router')
            }
            self.Routers.append(dict)
        
        for Service in root.iter('Service'): #root.findall('Service'):
            dict = {
                'type':Service.get('type'),
                'uuid':Service.get('uuid'),
                'name':Service.get('name'),
                'systemid':Service.get('systemid'),
                'msid':Service.get('msid'),
                'server':Service.get('server'),
                'routerid':Service.get('routerid'),
                'dcpg':Service.get('dcpg'),
                'sncname':Service.get('sncname'),
            }
            if dict['type'] == 'SAPGUI':
                self.Services.append(dict)
            
        for MSService in root.iter('Messageserver'): # root.findall('Messageserver'):
            dict = {
                'uuid':MSService.get('uuid'),
                'name':MSService.get('name'),
                'description':MSService.get('description'),
                'host':MSService.get('host'),
                'port':MSService.get('port'),
            }
            self.Messageservers.append(dict)
     
        for service in self.Services:
            dict = {
                'name':service.get('name'),
                'system':service.get('systemid'),
                'systemNo':'' 
            }
            server = service.get('server')
            if server :
                serveraddress = server.split(':')
            dict['serverGroup']  = serveraddress[0]
            if len(serveraddress) > 1:
                dict['systemNo'] = serveraddress[1][2:4]
            if service['msid'] in [dict1.get('uuid') for dict1 in self.Messageservers ]:
                dict1 = sorted( self.Messageservers,key = lambda dicts:dicts['uuid'] != service['msid'])[0]
                dict['messageServer'] = dict1.get('host')
            
            if service['routerid'] in [dict1.get('uuid') for dict1 in self.Routers ]:
                dict1 = sorted( self.Routers,key = lambda dicts:dicts['uuid'] != service['routerid'])[0]
                dict['router'] = dict1.get('router')
            
            dict['dcpg'] = service.get('dcpg')
            dict['sncname']   = service.get('sncname')
            self.ServerList.append(dict) 
            
           
        # df1 = pandas.DataFrame(self.ServerList)
        # df1.to_excel('saspconfig.xlsx',index=False)

    
    
if __name__ == '__main__':
    pass
