******************WinExec**************************
#DEFINE SW_HIDE             0
#DEFINE SW_SHOWNORMAL       1
#DEFINE SW_NORMAL           1
#DEFINE SW_SHOWMINIMIZED    2
#DEFINE SW_SHOWMAXIMIZED    3
#DEFINE SW_MAXIMIZE         3
#DEFINE SW_SHOWNOACTIVATE   4
#DEFINE SW_SHOW             5
#DEFINE SW_MINIMIZE         6
#DEFINE SW_SHOWMINNOACTIVE  7
#DEFINE SW_SHOWNA           8
#DEFINE SW_RESTORE          9
#DEFINE SW_SHOWDEFAULT      10
#DEFINE SW_MAX              10
******************WinExec**************************


PROCEDURE controlerror
PARAMETERS merror, MESS, mess1, mprog, mlineno, oMessage AS LABEL
PRIVATE iForReading,iForWriting,iForAppending
PRIVATE oFs AS OBJECT
PRIVATE oTx AS OBJECT
PRIVATE sCa AS STRING
PRIVATE sFe AS STRING
iForReading = 1
iForWriting = 2
iForAppending = 8
WAIT WINDOW "Error: "+MESS TIMEOUT 1
sCa = 'Error number: ' + LTRIM(STR(merror))+ " - "+;
	'Error message: ' + MESS+" - "+;
	'Line of code with error: ' + mess1+" - "+;
	'Line number of error: ' + LTRIM(STR(mlineno))+" - "+;
	'Program with error: ' + mprog

sFe =".\History.Err"
oFs = CREATEOBJECT("Scripting.FileSystemObject")
IF oFs.FileExists(sFe) = .F. THEN
	oTx = oFs.CreateTextFile(sFe, .T.)
ELSE
	oTx = oFs.OpenTextFile(sFe, iForAppending, -2)
ENDIF
oTx.WriteLine(sCa)
oTx.CLOSE
ENDPROC

FUNCTION APPalreadyrunning
LOCAL hsem, lpszSemName
#DEFINE ERROR_ALREADY_EXISTS 183
DECLARE INTEGER GetLastError IN win32API
DECLARE INTEGER CreateSemaphore IN WIN32API ;
	STRING @ lpSemaphoreAttributes, ;
	LONG lInitialCount, ;
	LONG lMaximumCount, ;
	STRING @ lpName
lpszSemName = "CadenaUnicadetuAplicacion"
hsem = CreateSemaphore(0,0,1,lpszSemName)
RETURN (hsem # 0 AND GetLastError() == ERROR_ALREADY_EXISTS)
ENDFUNC

FUNCTION Scramble(sTEXT AS STRING) AS STRING
PRIVATE i AS INTEGER
PRIVATE c AS INTEGER
PRIVATE sTemp AS STRING
sTemp = ""
FOR i = 1 TO LEN(sTEXT)
	c = ASC(SUBSTR(sTEXT, i, 1))
	c = c + 10

	IF c > 255 THEN
		c = c - 255
	ENDIF
	sTemp = sTemp + CHR(c)
NEXT i
RETURN sTemp
ENDFUNC

FUNCTION UnScramble(sTEXT AS STRING) AS STRING
PRIVATE i AS INTEGER
PRIVATE c AS INTEGER
PRIVATE sTemp AS STRING
sTemp = ""
FOR i = 1 TO LEN(sTEXT)
	c = ASC(SUBSTR(sTEXT, i, 1))
	c = c - 10
	IF c < 0 THEN
		c = 256 + c
	ENDIF
	sTemp = sTemp + CHR(c)
NEXT i
RETURN sTemp
ENDFUNC

FUNCTION Read_Ini_File(cNombreFichero AS STRING, cSeccion AS STRING, cClave AS STRING, cCadenaRetorno) AS STRING
PRIVATE LTmp AS LONG
#DEFINE MAX_SECTION 255
DECLARE ;
	INTEGER GetPrivateProfileString IN WIN32API;
	STRING   cSeccion, ;
	STRING   cClave, ;
	STRING   cDefecto, ;
	STRING  @cCadenaRetorno, ;
	INTEGER  nTama, ;
	STRING   cNombreFichero
nBufferSize=255
cCadenaRetorno = REPLICATE( CHR(0), MAX_SECTION )
LTmp =GetPrivateProfileString(cSeccion, cClave, "none",@cCadenaRetorno, nBufferSize, cNombreFichero)
IF LTmp = 0 THEN
	RETURN ""
ELSE
	RETURN LEFT(cCadenaRetorno, LTmp)
ENDIF
ENDFUNC

PROCEDURE Write_Ini_File(lpFileName AS STRING, lpAppName AS STRING, lpKeyName AS STRING, lpString AS STRING) AS STRING
PRIVATE LTmp AS LONG
DECLARE LONG WritePrivateProfileString IN kernel32 ;
	STRING lpApplicationName,STRING lpKeyName, STRING lpString, STRING lpFileName
LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
ENDPROC

FUNCTION GetSystemDirectory
DECLARE INTEGER GetSystemDirectory;
	IN kernel32 AS GetSystemDirectoryA;
	STRING @ lpBuffer,;
	INTEGER nSize
LOCAL cSystemPath, lnLength
cSystemPath = SPACE(1024)
lnLength = GetSystemDirectoryA(@cSystemPath,1023)
CLEAR DLLS GetSystemDirectoryA
RETURN LEFT(cSystemPath,lnLength)
ENDFUNC

FUNC GetHDSerial(tcRootpath)
DECLARE GetVolumeInformation IN win32api STRING, STRING @, ;
	INTEGER, INTEGER @, INTEGER @, INTEGER @, STRING @, INTEGER

lcName = SPACE(255)
lnNameLen = LEN(lcName)
lnSerialNumber = 0
lnFilenameMaxLength = 0
lnFileSystemFlags = 0
lcFileSystemName = SPACE(255)
lnFSNameLen = LEN(lcFileSystemName)

=GetVolumeInformation(tcRootpath, @lcName, lnNameLen, ;
	@lnSerialNumber, @lnFilenameMaxLength, ;
	@lnFileSystemFlags, @lcFileSystemName, lnFSNameLen)
IF EMPTY(lnSerialNumber)
	RETURN ""
ELSE
	IF lnSerialNumber < 0
		lnSerialNumber = (2^32) + lnSerialNumber
	ENDIF
	lcSerial = TRANSFORM(lnSerialNumber,"@0")
	lcSerial = TRANSFORM(SUBSTR(lcSerial,3),"@R XXXX-XXXX")
	RETURN lcSerial
ENDIF
ENDFUNC

FUNCTION Get_CPU_Id
LOCAL lcComputerName, loWMI, lowmiWin32Objects, lowmiWin32Object,sOut
lcComputerName = GETWORDNUM(SYS(0),1)
loWMI = GETOBJECT("WinMgmts://" + lcComputerName)
lowmiWin32Objects = loWMI.InstancesOf("Win32_Processor")
sOut = ""
FOR EACH lowmiWin32Object IN lowmiWin32Objects
	WITH lowmiWin32Object
		sOut = sOut +TRANSFORM(.ProcessorID)
	ENDWITH
ENDFOR
RETURN sOut
ENDFUNC

FUNCTION Get_MAC_Address
LOCAL lcComputerName, loWMIService, loItems, loItem, lcMACAddress
lcComputerName = "."
loWMIService = GETOBJECT("winmgmts:\\" + lcComputerName + "\root\cimv2")
loItems = loWMIService.ExecQuery("Select * from Win32_NetworkAdapter",,48)
FOR EACH loItem IN loItems
	lcMACAddress = loItem.MACAddress
	IF !ISNULL(lcMACAddress)
		? "MAC Address: " + loItem.MACAddress
		?
	ENDIF
ENDFOR
ENDFUNC

FUNCTION Get_Volume_Serial_Numbers
LOCAL lcComputerName, loWMIService, loItems, loItem, lcVolumeSerial
lcComputerName = "."
loWMIService = GETOBJECT("winmgmts:\\" + lcComputerName + "\root\cimv2")
loItems = loWMIService.ExecQuery("Select * from Win32_LogicalDisk")
FOR EACH loItem IN loItems
	lcVolumeSerial = loItem.VolumeSerialNumber
	IF !ISNULL(lcVolumeSerial)
		? "Name: " + loItem.NAME
		? "Volume Serial Number: " + loItem.VolumeSerialNumber
		?
	ENDIF
ENDFOR
ENDFUNC

PROCEDURE GetMotherBoardNumber()
LOCAL loloc, lowmi, locolboard, loboard, loproperty, lcNumber, x, lnNumber
loloc = CREATEOBJECT('WbemScripting.SWbemLocator')
lowmi = loloc.ConnectServer()
locolboard = lowmi.InstancesOf('Win32_Baseboard')
FOR EACH loboard IN locolboard
	FOR EACH loproperty IN loboard.Properties_
		IF INLIST(UPPER(loproperty.NAME),'SERIALNUMBER') THEN
			lcNumber = loproperty.VALUE
			EXIT
		ENDIF
		loproperty = .NULL.
	ENDFOR
	loboard = .NULL.
	IF NOT EMPTY(lcNumber)
		EXIT
	ENDIF
ENDFOR
STORE .NULL. TO lowmi, loloc

IF NOT EMPTY(lcNumber)
	lnNumber = 1
	FOR x=1 TO LEN(lcNumber)
		lnNumber = lnNumber * ASC(SUBSTR(lcNumber,x,1))
	NEXT
	lnNumber = VAL(LEFT(STRTRAN(TRANSFORM(lnNumber),'.'),8))
ELSE
	lnNumber = ABS(THIS.GetVolumeNumber(ADDBS(JUSTDRIVE(GETENV("windir")))))
ENDIF
RETURN lnNumber
ENDPROC

PROCEDURE Set_Tmp_DBF
CLOSE DATABASES
CLOSE TABLES
COPY FILE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file01.dbf") TO (sgDir_Tmp+"\file01.dbf")
COPY FILE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file02.dbf") TO (sgDir_Tmp+"\file02.dbf")
COPY FILE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file03.dbf") TO (sgDir_Tmp+"\file03.dbf")
COPY FILE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file05.dbf") TO (sgDir_Tmp+"\file05.dbf")
ENDPROC

FUNCTION  Conectar_DBPro AS Boolean
IF !USED("file01")
	USE (sgDir_Tmp+"\file01") IN 0 SHARED
ENDIF
IF !USED("file02")
	USE (sgDir_Tmp+"\file02") IN 0 SHARED
ENDIF
IF !USED("file03")
	USE (sgDir_Tmp+"\file03") IN 0 SHARED
ENDIF
IF !USED("file05")
	USE (sgDir_Tmp+"\file05") IN 0 SHARED
ENDIF
SELECT ;
	file01.ID_GEN, file01.ID_ORD AS ID_ORD1, ;
	file02.ID_DIS ,file02.ID_ORD AS ID_ORD2, ;
	file02.FL_IMG, ;
	file03.* FROM file01,file02,file03 ;
	WHERE  file01.ID_GEN=file02.ID_GEN ;
	AND    file03.ID_DIS=file02.ID_DIS ;
	INTO TABLE (sgDir_Tmp+"\QRYPUB")
*	HAVING file03.FL_PRC = 1;

IF !USED("QRYPUB")
	USE (sgDir_Tmp+"\QRYPUB") IN 0 SHARED
ENDIF
SELECT QRYPUB
IF RECCOUNT()>0 THEN
	RETURN .T.
ELSE
	RETURN .F.
ENDIF
ENDPROC

FUNCTION SPLIT
PARAMETERS sValue AS STRING, sDelim AS STRING ,aDatos AS STRING
LOCAL iFn AS INTEGER, iCnt AS INTEGER, i AS INTEGER
DIMENSION aDatos(1) AS STRING
iCnt=0
i=0
iCnt=OCCURS(sDelim,sValue)+1
DIMENSION aDatos(iCnt)
FOR i=1 TO iCnt
	iFn=AT(sDelim,sValue)
	aDatos(i)=SUBSTR(sValue,1,IIF(iFn=0,LEN(sValue),iFn-1))
	sValue=SUBSTR(sValue,iFn+1,LEN(sValue))
NEXT i
RETURN iCnt
ENDFUNC

FUNCTION ranColor
RETURN RGB((INT(255 * RAND(-1)) + 1),(INT(255 * RAND(-1)) + 1),(INT(255 * RAND(-1)) + 1))
ENDFUNC

PROCEDURE Borra_Ref_Can
PARAMETER bAll_Gen AS Boolean,pGen AS INTEGER, pDis AS INTEGER
PRIVATE bUsed1 AS Boolean, bUsed2 AS Boolean, sFile1 AS STRING
PRIVATE sCadena AS STRING , iTot AS INTEGER, iCount AS INTEGER
PRIVATE oFs AS OBJECT
PRIVATE sCa AS STRING
STORE .F. TO bUsed1,bUsed2
oFs = CREATEOBJECT("Scripting.FileSystemObject")
IF USED("file02") THEN
	USE IN file02
	bUsed1=.T.
ENDIF
IF USED("file03") THEN
	USE IN file03
	bUsed2=.T.
ENDIF
USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file02") IN 0 EXCLU
USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file03") IN 0 EXCLU
SELECT file03
INDEX ON ID_DIS TAG ID_DIS
SET ORDER TO TAG ID_DIS
SELECT file02
INDEX ON ID_DIS TAG ID_DIS
SET ORDER TO TAG ID_DIS
SELECT file03
IF bAll_Gen=.T. THEN
	SET RELATION TO
	SET RELATION TO file03.ID_DIS INTO file02 ADDITIVE
	GO TOP
	REPLACE file03.MARK WITH "X" FOR file02.ID_GEN=pGen
ELSE
	REPLACE file03.MARK WITH "X" FOR file03.ID_DIS=pDis
ENDIF
SELECT file03
SET RELATION TO
COUNT TO iTot FOR !EMPTY(file03.MARK)
GO TOP
iCount=0
=Write_TO_Histd(oFs,TTOC(DATETIME()))
sCa="------------------BORRANDO REGISTRO DE CANCIONES----------------"
=Write_TO_Histd(oFs,sCa)
SCAN FOR !EMPTY(file03.MARK)
	iCount=iCount+1
	sFile1=UPPER(ALLTRIM(FORCEPATH(file03.fl_mp3,sgDir_Mp3)))
	IF FILE(sFile1)
		sCadena="("+ALLTRIM(STR(iCount))+" de "+ALLTRIM(STR(iTot))+"):. "
		WAIT WINDOW "Borrando archivo "+sCadena+"["+sFile1+"]" NOWAIT
		DELETE FILE (sFile1)
		=Write_TO_Histd(oFs,"Borrando archivo "+sCadena+"["+sFile1+"]")
	ELSE
		WAIT WINDOW "["+sFile1+"] no encontrado..." NOWAIT
		=Write_TO_Histd(oFs,"["+sFile1+"] no encontrado...")
	ENDIF
ENDSCAN
=Write_TO_Histd(oFs,"*FIN*")
SELECT file03
DELETE FOR !EMPTY(file03.MARK)
PACK
SELECT file02
DELETE TAG ALL
USE IN file02
IF bUsed1=.T. THEN
	USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file02") IN 0 SHARED
ENDIF
SELECT file03
DELETE TAG ALL
USE IN file03
IF bUsed2=.T. THEN
	USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file03") IN 0 SHARED
ENDIF
RELEASE bUsed1, bUsed2 , sFile1,sCadena,iTot,iCount
RELEASE oFs,sCa
ENDPROC

PROCEDURE Borra_Ref_Dis
PARAMETER bAll_Gen AS Boolean,pGen AS INTEGER, pDis AS INTEGER
PRIVATE bUsed1 AS Boolean, bUsed2 AS Boolean,  sFile1 AS STRING
PRIVATE sCadena AS STRING , iTot AS INTEGER, iCount AS INTEGER
PRIVATE oFs AS OBJECT
PRIVATE sCa AS STRING
STORE .F. TO bUsed1,bUsed2
oFs = CREATEOBJECT("Scripting.FileSystemObject")
IF USED("file02") THEN
	USE IN file02
	bUsed1=.T.
ENDIF
USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file02") IN 0 EXCLU
SELECT file02
INDEX ON ID_DIS TAG ID_DIS
SET ORDER TO TAG ID_DIS
SET RELATION TO
GO TOP
IF bAll_Gen =.T. THEN
	REPLACE;
		file02.MARK WITH "X" FOR file02.ID_GEN=pGen
ELSE
	REPLACE;
		file02.MARK WITH "X" FOR file02.ID_DIS=pDis
ENDIF
COUNT TO iTot FOR !EMPTY(file02.MARK)
GO TOP
iCount=0
=Write_TO_Histd(oFs,TTOC(DATETIME()))
sCa="------------------BORRANDO REGISTRO DE DISCOS--------------------"
=Write_TO_Histd(oFs,sCa)
SCAN FOR !EMPTY(file02.MARK)
	iCount=iCount+1
	sFile1=UPPER(ALLTRIM(FORCEPATH(file02.FL_IMG,sgDir_Img)))
	IF FILE(sFile1)
		sCadena="("+ALLTRIM(STR(iCount))+" de "+ALLTRIM(STR(iTot))+"):. "
		WAIT WINDOW "Borrando car�tula "+sCadena+"["+sFile1+"]" NOWAIT
		DELETE FILE (sFile1)
		=Write_TO_Histd(oFs,"Borrando car�tula "+sCadena+"["+sFile1+"]")
	ELSE
		WAIT WINDOW "["+sFile1+"] no encontrado..." NOWAIT
		=Write_TO_Histd(oFs,"["+sFile1+"] no encontrado...")
	ENDIF
ENDSCAN
=Write_TO_Histd(oFs,"*FIN*")
SELECT file02
DELETE FOR !EMPTY(file02.MARK)
PACK
DELETE TAG ALL
USE IN file02
IF bUsed1=.T. THEN
	USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file02") IN 0 SHARED
ENDIF
RELEASE bUsed1, bUsed2 , sFile1,sCadena,iTot,iCount
RELEASE oFs,sCa
ENDPROC

PROCEDURE Borra_Ref_Gen
PARAMETERS ipGen AS INTEGER
PRIVATE bUsed1 AS Boolean, sFile1 AS STRING
PRIVATE sCadena AS STRING , iTot AS INTEGER, iCount AS INTEGER
IF USED("file01") THEN
	USE IN file01
	bUsed1=.T.
ENDIF
USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file01") IN 0 EXCLU
SELECT file01
DELETE FOR ID_GEN=ipGen
PACK
DELETE TAG ALL
USE IN file01
IF bUsed1=.T. THEN
	USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file01") IN 0 SHARED
ENDIF
RELEASE bUsed1, sFile1,sCadena,iTot,iCount
ENDPROC

FUNCTION Write_TO_Histd
PARAMETERS poFS AS OBJECT,psCa AS STRING
PRIVATE iForReading,iForWriting,iForAppending
PRIVATE oTx AS OBJECT
PRIVATE sFe AS STRING
iForReading = 1
iForWriting = 2
iForAppending = 8
sFe =CURDIR() + "\HISTORY.del"
IF oFs.FileExists(sFe) = .F. THEN
	oTx = oFs.CreateTextFile(sFe, .T.)
ELSE
	oTx = oFs.OpenTextFile(sFe, iForAppending, -2)
ENDIF
oTx.WriteLine(psCa)
oTx.CLOSE
ENDPROC

PROCEDURE check_integ_01
PARAMETERS psPath
PRIVATE bUsed AS Boolean, bFound1 AS Boolean, bFound2 AS Boolean, bFound3 AS Boolean, afls AS STRING, X AS INTEGER
PRIVATE slDir_Fls AS STRING
IF EMPTY(psPath) THEN
	slDir_Fls=IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)
ELSE
	slDir_Fls=psPath
ENDIF
STORE .F. TO bUsed,bFound1, bFound2, bFound3
STORE 0 TO X
WAIT WINDOW "CHEKEANDO INTEGRIDAD DE [FILE01]..." NOWAIT
IF  USED("file01")
	SELECT file01
	USE
	bUsed=.T.
ENDIF
USE (slDir_Fls+"\file01") IN 0 EXCLUSIVE
SELECT file01
DELETE TAG ALL
=AFIELDS(afls)
FOR X=1 TO ALEN(afls,1)
	IF UPPER(afls(X,1))="G_KAR" THEN
		bFound1=.T.
	ENDIF
	IF UPPER(afls(X,1))="ULT_ACT" THEN
		bFound2=.T.
	ENDIF
	IF UPPER(afls(X,1))="MARK" THEN
		bFound3=.T.
	ENDIF
NEXT
IF bFound1=.T.
	ALTER TABLE (slDir_Fls +"\file01") DROP COLUMN G_KAR
ENDIF
IF bFound2=.F.
	ALTER TABLE (slDir_Fls +"\file01") ADD COLUMN ULT_ACT T(8)
	SELECT file01
	REPLACE ULT_ACT WITH DATETIME()	ALL
	GO TOP
ENDIF
IF bFound3=.F.
	ALTER TABLE (slDir_Fls +"\file01") ADD COLUMN MARK c(1)
	SELECT file01
	REPLACE MARK WITH "" ALL
	GO TOP
ENDIF
SELECT file01
USE
IF 	bUsed=.T. THEN
	USE (slDir_Fls +"\file01") IN 0 SHARED
ENDIF
RELEASE bUsed,bFound1, bFound2,afls, X
ENDPROC

PROCEDURE check_integ_02
PARAMETERS psPath
PRIVATE slDir_Fls AS STRING
PRIVATE bUsed AS Boolean,;
	bFound1 AS Boolean, bFound2 AS Boolean,bFound3 AS Boolean, bFound4 AS Boolean, bFound5 AS Boolean, bFound6 AS Boolean,;
	bFound7 AS Boolean,afls AS STRING, X AS INTEGER
IF EMPTY(psPath) THEN
	slDir_Fls=IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)
ELSE
	slDir_Fls=psPath
ENDIF
STORE .F. TO bUsed,bFound1, bFound2,bFound3,bFound4,bFound5,bFound6,bFound7
STORE 0 TO X
IF  USED("file02")
	SELECT file02
	USE
	bUsed=.T.
ENDIF
USE (slDir_Fls +"\file02") IN 0 EXCLUSIVE
SELECT file02
DELETE TAG ALL
=AFIELDS(afls)
FOR X=1 TO ALEN(afls,1)
	IF UPPER(afls(X,1))="C_VIDEO" THEN
		bFound1=.T.
	ENDIF
	IF UPPER(afls(X,1))="ULT_ACT" THEN
		bFound2=.T.
	ENDIF
	IF UPPER(afls(X,1))="D_PROMO" THEN
		bFound3=.T.
	ENDIF
	IF UPPER(afls(X,1))="D_KAR" THEN
		bFound4=.T.
	ENDIF
	IF UPPER(afls(X,1))="FL_NEW" THEN
		bFound5=.T.
	ENDIF
	IF UPPER(afls(X,1))="FL_PRD" THEN
		bFound6=.T.
	ENDIF
	IF UPPER(afls(X,1))="COUNTER" THEN
		bFound7=.T.
	ENDIF
NEXT
IF bFound1=.F.
	ALTER TABLE (slDir_Fls +"\file02") ADD COLUMN C_VIDEO  N(1)
	SELECT file02
	REPLACE C_VIDEO WITH 0 ALL
	GO TOP
ENDIF
IF bFound2=.F.
	ALTER TABLE (slDir_Fls +"\file02") ADD COLUMN ULT_ACT T(8)
	SELECT file02
	REPLACE ULT_ACT WITH DATETIME()	ALL
	GO TOP
ENDIF
*IF bFound3=.F.
*	ALTER TABLE (slDir_Fls +"\file02") ADD COLUMN D_PROMO N(1)
*	SELECT file02
*	REPLACE D_PROMO WITH 0	ALL
*	GO TOP
*ENDIF
IF bFound3=.T.
	ALTER TABLE (slDir_Fls +"\file02") DROP COLUMN D_PROMO
ENDIF
*IF bFound4=.F.
*	ALTER TABLE (slDir_Fls +"\file02") ADD COLUMN D_KAR N(1)
*	SELECT file02
*	REPLACE D_KAR WITH 0 ALL
*	GO TOP
*ENDIF
IF bFound4=.T.
	ALTER TABLE (slDir_Fls +"\file02") DROP COLUMN D_KAR
ENDIF
IF bFound5=.F.
	ALTER TABLE (slDir_Fls +"\file02") ADD COLUMN FL_NEW N(1)
	SELECT file02
	REPLACE FL_NEW WITH 0 ALL
	GO TOP
ENDIF
IF bFound6=.F.
	ALTER TABLE (slDir_Fls +"\file02") ADD  COLUMN FL_PRD N(1)
	SELECT file02
	REPLACE FL_PRD WITH 0 ALL
	GO TOP
ENDIF
IF bFound7=.F.
	ALTER TABLE (slDir_Fls +"\file02") ADD  COLUMN COUNTER B(8)
	SELECT file02
	REPLACE COUNTER WITH 0 ALL
	GO TOP
ENDIF
SELECT file02
USE
IF 	bUsed=.T. THEN
	USE (slDir_Fls +"\file02") IN 0 SHARED
ENDIF
RELEASE bUsed,bFound1, bFound2, bFound3, bFound4, bFound5, bFound6,afls,X
ENDPROC

PROCEDURE check_integ_03
PARAMETERS psPath
PRIVATE slDir_Fls AS STRING
PRIVATE bUsed AS Boolean,;
	bFound1 AS Boolean, bFound2 AS Boolean,bFound3 AS Boolean, bFound4 AS Boolean, bFound5 AS Boolean, bFound6 AS Boolean,;
	bFound7 AS Boolean, bFound8 AS Boolean, bFound9 AS Boolean, X AS INTEGER
IF EMPTY(psPath) THEN
	slDir_Fls=IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)
ELSE
	slDir_Fls=psPath
ENDIF
STORE .F. TO bUsed,bFound1, bFound2,bFound3,bFound4,bFound5,bFound6,bFound7,bFound8,bFound9
STORE 0 TO X
IF  USED("file03")
	SELECT file03
	USE
	bUsed=.T.
ENDIF
USE (slDir_Fls +"\file03") IN 0 EXCLUSIVE
SELECT file03
DELETE TAG ALL
=AFIELDS(afls)
FOR X=1 TO ALEN(afls,1)
	IF UPPER(afls(X,1))="C_KAR" THEN
		bFound1=.T.
	ENDIF
	IF UPPER(afls(X,1))="ULT_ACT" THEN
		bFound2=.T.
	ENDIF
	IF UPPER(afls(X,1))="C_PROMO" THEN
		bFound3=.T.
	ENDIF
	IF UPPER(afls(X,1))="FL_PRC" THEN
		bFound4=.T.
	ENDIF
	IF UPPER(afls(X,1))="FL_KAG" THEN
		bFound5=.T.
	ENDIF
	IF UPPER(afls(X,1))="ID_GEN" THEN
		bFound6=.T.
	ENDIF
	IF UPPER(afls(X,1))="COUNTER" THEN
		bFound7=.T.
	ENDIF
	IF UPPER(afls(X,1))="C_VIDEO" THEN
		bFound8=.T.
	ENDIF
	IF UPPER(afls(X,1))="FLE_SEC" THEN
		bFound9=.T.
	ENDIF
NEXT
IF bFound1=.T. THEN
	ALTER TABLE (slDir_Fls +"\file03") DROP COLUMN C_KAR
ENDIF
IF bFound2=.F.
	ALTER TABLE (slDir_Fls +"\file03") ADD COLUMN ULT_ACT T(8)
	SELECT file03
	REPLACE ULT_ACT WITH DATETIME()	ALL
	GO TOP
ENDIF
IF bFound3=.T. THEN
	ALTER TABLE (slDir_Fls +"\file03") DROP COLUMN C_PROMO
ENDIF
IF bFound4=.F. THEN
	ALTER TABLE (slDir_Fls +"\file03") ADD COLUMN FL_PRC N(1)
	SELECT file03
	REPLACE FL_PRC WITH 0 ALL
	GO TOP
ENDIF
IF bFound5=.T. THEN
	ALTER TABLE (slDir_Fls +"\file03") DROP COLUMN FL_KAG
ENDIF
IF bFound6=.T. THEN
	ALTER TABLE (slDir_Fls +"\file03") DROP COLUMN ID_GEN
ENDIF
IF bFound7=.F. THEN
	ALTER TABLE (slDir_Fls +"\file03") ADD COLUMN COUNTER B(8)
	SELECT file03
	REPLACE COUNTER WITH 0 ALL
	GO TOP
ENDIF
IF bFound8=.F. THEN
	ALTER TABLE (slDir_Fls +"\file03") ADD COLUMN C_VIDEO c(1)
	GO TOP
ENDIF
IF bFound9=.F. THEN
	ALTER TABLE (slDir_Fls +"\file03") ADD COLUMN FLE_SEC N(15)
	SELECT file03
	GO TOP
	REPLACE;
		FLE_SEC WITH INT(VAL(LEFT(ALLTRIM(JUSTFNAME(fl_mp3)),LEN(ALLTRIM(JUSTFNAME(fl_mp3)))-4))) ALL
	GO TOP
ENDIF
SELECT file03
USE
IF 	bUsed=.T. THEN
	USE (slDir_Fls +"\file03") IN 0 SHARED
ENDIF
RELEASE bUsed,bFound1, bFound2, bFound3, bFound4, bFound5, bFound6,bFound7,bFound8,bFound9,afls,X
ENDPROC

PROCEDURE check_integ_05
PARAMETERS psPath
PRIVATE slDir_Fls AS STRING
PRIVATE bUsed AS Boolean,;
	bFound1 AS Boolean, bFound2 AS Boolean,bFound3 AS Boolean, bFound4 AS Boolean,;
	fls AS STRING, X AS INTEGER
IF EMPTY(psPath) THEN
	slDir_Fls=IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)
ELSE
	slDir_Fls=psPath
ENDIF
STORE .F. TO bUsed,bFound1, bFound2,bFound3,bFound4
STORE 0 TO X
IF  USED("file05")
	SELECT file05
	USE
	bUsed=.T.
ENDIF
USE (slDir_Fls +"\file05") IN 0 EXCLUSIVE
SELECT file05
DELETE TAG ALL
=AFIELDS(afls)
FOR X=1 TO ALEN(afls,1)
	IF UPPER(afls(X,1))="VIP" THEN
		bFound1=.T.
	ENDIF
	IF UPPER(afls(X,1))="VIDEO" THEN
		bFound2=.T.
	ENDIF
	IF UPPER(afls(X,1))="ID_TIPO" THEN
		bFound3=.T.
	ENDIF
	IF UPPER(afls(X,1))="ID_ORD" THEN
		bFound4=.T.
	ENDIF
NEXT
IF bFound1=.F.
	ALTER TABLE (slDir_Fls +"\file05") ADD COLUMN VIP c(4)
ENDIF
IF bFound2=.F.
	ALTER TABLE (slDir_Fls +"\file05") ADD COLUMN VIDEO c(4)
ENDIF
IF bFound3=.F.
	ALTER TABLE (slDir_Fls +"\file05") ADD COLUMN ID_TIPO c(4)
ENDIF
IF bFound4=.F. THEN
	ALTER TABLE (slDir_Fls +"\file05") ADD COLUMN ID_ORD c(10)
ENDIF
SELECT file05
*DELETE ALL
PACK
USE
IF 	bUsed=.T. THEN
	USE (slDir_Fls +"\file05") IN 0 SHARED
ENDIF
RELEASE bUsed,bFound1, bFound2, bFound3, bFound4,X
ENDPROC

FUNCTION DO_WinExec
PARAMETERS cCmdLine AS STRING,nCmdShow AS INTEGER
DECLARE ;
	INTEGER WinExec ;
	IN WIN32API ;
	STRING   cCmdLine, ;
	INTEGER  nCmdShow
=WinExec(cCmdLine, nCmdShow)
ENDFUNC


*----------------------------------------
* FUNCTION GetConnection(lcDrive)
*----------------------------------------
* Retorna el nombre de la PC y recurso
* compartido de una conexi�n de red
* PARAMETROS: lcDrive
* USO: ? GetConnection("K:")
*----------------------------------------
FUNCTION GetConnection(lcDrive)
DECLARE INTEGER WNetGetConnection IN WIN32API ;
	STRING lpLocalName, ;
	STRING @lpRemoteName, ;
	INTEGER @lpnLength
LOCAL cRemoteName, nLength, lcRet, llRet
cRemoteName=SPACE(100)
nLength = 100
llRet = WNetGetConnection(lcDrive,@cRemoteName,@nLength)
lcRet = LEFT(cRemoteName,AT(CHR(0),cRemoteName)-1)
RETURN lcRet
ENDFUNC

*----------------------------------------
* FUNCTION AddConnection(tcDrive,tcResource,tcPassword)
*----------------------------------------
* Conecta un recurso compartido a la unidad tcDrive
* USO: ? AddConnection("Z:","PC_REMOTARECURSO")
*----------------------------------------
FUNCTION AddConnection(tcDrive,tcResource,tcPassword)
LOCAL lnRet
DECLARE INTEGER WNetAddConnection IN WIN32API;
	STRING @lpzRemoteName, ;
	STRING @lpzPassword,;
	STRING @lpzLocalName
IF PARAMETERS() < 3
	lnRet = WNetAddConnection(@tcResource,0,@tcDrive)
ELSE
	lnRet = WNetAddConnection(@tcResource,@tcPassword, @tcDrive)
ENDIF
IF lnRet # 0
	RETURN "Error " + ALLT(STR(lnRet)) + ;
		" al conectar el drive " + tcDrive
ENDIF
RETURN ""
ENDFUNC

*----------------------------------------
* FUNCTION CancelConnection(tcDrive)
*----------------------------------------
* Desconecta una unidad de red
* USO: ? CancelConnection("Z:")
*----------------------------------------
FUNCTION CancelConnection(tcDrive)
LOCAL lnRet
DECLARE INTEGER WNetCancelConnection IN WIN32API;
	STRING @lpzLocalName, ;
	INTEGER nForce
lnRet = WNetCancelConnection( @tcDrive, 0)
IF lnRet # 0
	RETURN "Error " + ALLT(STR(lnRet)) + ;
		" al desconectar el drive " + tcDrive
ENDIF
RETURN ""
ENDFUNC

*----------------------------------------
* FUNCTION WinDir()
*----------------------------------------
* Retorna el directorio de Windows
* USO: ? WinDir() -> "C:WINNT"
*----------------------------------------
FUNCTION WinDir()
LOCAL lcPath, lnSize
lcPath = SPACE(255)
lnSize = 255
DECLARE INTEGER GetWindowsDirectory IN Win32API ;
	STRING @pszSysPath,;
	INTEGER cchSysPath
lnSize = GetWindowsDirectory(@lcPath, lnSize)
IF lnSize <= 0
	lcPath = ""
ELSE
	lcPath = ADDBS(SUBSTR(lcPath, 1, lnSize))
ENDIF
RETURN lcPath
ENDFUNC

*----------------------------------------
* FUNCTION SystemDir()
*----------------------------------------
* Retorna el directorio SYSTEM de Windows
* USO: ? SystemDir() -> "C:WINNTSYSTEM32"
*----------------------------------------
FUNCTION SystemDir()
LOCAL lcPath, lnSize
lcPath = SPACE(255)
lnSize = 255
DECLARE INTEGER GetSystemDirectory IN Win32API ;
	STRING @pszSysPath,;
	INTEGER cchSysPath
lnSize = GetSystemDirectory(@lcPath, lnSize)
IF lnSize <= 0
	lcPath = ""
ELSE
	lcPath = ADDBS(SUBSTR(lcPath, 1, lnSize))
ENDIF
RETURN lcPath
ENDFUNC

*----------------------------------------
* FUNCTION TempDir()
*----------------------------------------
* Retorna la ruta de los archivos temporales
* USO: ? TempDir() -> "C:WINNTTEMP"
*----------------------------------------
FUNCTION TempDir()
LOCAL lcPath, lnRet
lcPath = SPACE(255)
lnSize = 255
DECLARE INTEGER GetTempPath IN WIN32API ;
	INTEGER nBufSize, ;
	STRING @cPathName
lnRet = GetTempPath(lnSize, @lcPath)
IF lnRet <= 0
	lcPath = ""
ELSE
	lcPath = ADDBS(SUBSTR(lcPath, 1, lnRet))
ENDIF
RETURN lcPath
ENDFUNC


*----------------------------------------
* FUNCTION UserName()
*----------------------------------------
* Retorna el nombre del usuario
* USO: ? UserName() -> "LUISG"
*----------------------------------------
FUNCTION UserName()
LOCAL lcUser, lnSize
lcUser = SPACE(80)
lnSize = 80
DECLARE INTEGER GetUserName IN WIN32API ;
	STRING @cUserName, ;
	INTEGER @nSize
=GetUserName(@lcUser, @lnSize)
IF lnSize < 2
	lcUser = ""
ELSE
	lcUser = SUBSTR(lcUser, 1, lnSize-1)
ENDIF
RETURN lcUser
ENDFUNC

*----------------------------------------
* FUNCTION ComputerName()
*----------------------------------------
* Retorna el nombre de la computadora
* USO: ? ComputerName() -> "PC_DESARROLLO"
*----------------------------------------
FUNCTION ComputerName()
LOCAL lcComputer, lnSize
lcComputer = SPACE(80)
lnSize = 80
DECLARE INTEGER GetComputerName IN WIN32API ;
	STRING @cComputerName, ;
	INTEGER @nSize
=GetComputerName(@lcComputer, @lnSize)
IF lnSize < 2
	lcComputer = ""
ELSE
	lcComputer = SUBSTR(lcComputer, 1, lnSize)
ENDIF
RETURN lcComputer
ENDFUNC

*----------------------------------------
* FUNCTION Beep(tnSound)
*----------------------------------------
* Ejecuta el sonido predeterminado del sistema
* USO: ? Beep(0)
*----------------------------------------
FUNCTION Beep(tnSound)
tnSound = IIF(VARTYPE(tnSound) = "N", tnSound, 1)
DECLARE INTEGER MessageBeep IN WIN32API ;
	INTEGER nSound
RETURN IIF(MessageBeep(tnSound) = 1, .T., .F.)
ENDFUNC

*----------------------------------------
* FUNCTION PlayWav(lcWaveFile, lnPlayType)
*----------------------------------------
* Ejecuta un archivo .WAV
* USO: PlayWave( [,])
* Archivo_Wav = Ruta completa del archivo .Wav
* Ejecucion = 1 - Ejecuci�n en background (default)
* 0 - La aplicaci�n espera la ejecuci�n
* 2 - Si el archivo no existe, no ejecuta el default
* 4 - Apaga el sonido que se est� ejecutando
* 8 - Continuado
*----------------------------------------
FUNCTION PlayWav(lcWaveFile, lnPlayType)
lnPlayType = IIF(VARTYPE(lnPlayType) = "N", lnPlayType, 1)
DECLARE INTEGER PlaySound IN WINMM.DLL ;
	STRING cWave, ;
	INTEGER nModule, ;
	INTEGER nType
RETURN IIF(PlaySound(lcWaveFile,0,lnPlayType) = 1, .T., .F.)
ENDFUNC

*----------------------------------------
* FUNCTION Sleep(lnMiliSeg)
*----------------------------------------
* Funci�n que detiene la ejecuci�n de un programa
* durante "n" milisegundos sin interfase con el teclado.
* USO: ? Sleep(1500)
*----------------------------------------
FUNCTION Sleep(lnMiliSeg)
lnMiliSeg = IIF(VARTYPE(lnMiliSeg) = "N", lnMiliSeg, 1000)
DECLARE Sleep IN WIN32API ;
	INTEGER nMillisecs
RETURN IIF(Sleep(lnMiliSeg) = 1, .T., .F.)
ENDFUNC

*----------------------------------------
* FUNCTION SetCurPos(lnX, lnY)
*----------------------------------------
* Coloca el cursor en la posici�n especificada
* USO: ? SetCurPos(50,200)
*----------------------------------------
FUNCTION SetCurPos(lnX, lnY)
lnX = IIF(EMPTY(lnX),0,lnX)
lnY = IIF(EMPTY(lnY),0,lnY)
DECLARE INTEGER SetCursorPos IN WIN32API ;
	INTEGER lnX, ;
	INTEGER lnY
RETURN IIF(SetCursorPos(lnX, lnY) = 1, .T., .F.)
ENDFUNC

*----------------------------------------
* FUNCTION IsActive(tcCaption)
*----------------------------------------
* Verifica si una aplicaci�n ya est� activa
* USO: ? IsActive("Calculadora")
*----------------------------------------
FUNCTION IsActive(tcCaption)
DECLARE INTEGER FindWindow IN WIN32API ;
	STRING cNULL, ;
	STRING cWinName
RETURN FindWindow(0, tcCaption) # 0
ENDFUNC

*----------------------------------------
* FUNCTION YaActiva()
*----------------------------------------
* Comprueba que la aplicaci�n no se esta ejecutando
* Invoca a IsActive() descripta anteriormente
*----------------------------------------
FUNCTION YaActiva()
LOCAL llRet, lcCaption
llRet = .F.
lcCaption = _SCREEN.CAPTION
*--- Renombra temporariamente el caption de la app
_SCREEN.CAPTION = "_" + lcCaption
IF IsActive(lcCaption)
*--- Si ya esta activo
	MESSAGEBOX("Este sistema ya est� activo",16,"Aviso")
	llRet = .T.
ENDIF
_SCREEN.CAPTION = lcCaption
RETURN llRet
ENDFUNC

*----------------------------------------
* FUNCTION PPostMessage()
*----------------------------------------
* Comprueba que la aplicaci�n no se esta ejecutando
* Invoca a IsActive() descripta anteriormente
*----------------------------------------
PROCEDURE PostMessage
#DEFINE WM_CLOSE  16

DECLARE ;
	INTEGER PostMessage ;
	IN WIN32API ;
	INTEGER  nWnd, ;
	INTEGER  nMsg, ;
	INTEGER  nParam, ;
	INTEGER  nParam

*** Arranca en NotePad
WAIT WINDOW ;
	"Se va a arrancar el Bloc, " + ;
	"escriba algo en �l y espere..." ;
	TIMEOUT 3
RUN /N NotePad

*** Cierra preguntando si ha habido cambios
WAIT WINDOW ;
	"El Block se va ha cerrar..." ;
	TIMEOUT 5
*** La funci�n FindHwnd es un PRG
*nHwndNote = FindHwnd( "Bloc de notas" )
*=PostMessage( nHwndNote, WM_CLOSE, 0, 0 )
ENDPROC

PROCEDURE DEFINES_1
#DEFINE REG_CREATED_NEW_KEY 1
#DEFINE REG_OPENED_EXISTING_KEY 2
#DEFINE REG_OPTION_NON_VOLATILE 0
#DEFINE KEY_QUERY_VALUE 1
#DEFINE KEY_SET_VALUE 2
#DEFINE KEY_CREATE_SUB_KEY 4
#DEFINE KEY_ENUMERATE_SUB_KEYS 8
#DEFINE KEY_NOTIFY 16
#DEFINE KEY_CREATE_LINK 32
#DEFINE KEY_READ 1+8+16
#DEFINE KEY_WRITE 2+4
#DEFINE KEY_EXECUTE KEY_READ
#DEFINE KEY_ALL_ACCESS 1+2+4+8+16+32


#DEFINE HKEY_CLASSES_ROOT -2147483648
#DEFINE HKEY_CURRENT_USER -2147483647
#DEFINE HKEY_LOCAL_MACHINE -2147483646
#DEFINE HKEY_USERS -2147483645
#DEFINE HKEY_CURRENT_CONFIG -2147483653
#DEFINE HKEY_DYN_DATA -2147483654

#DEFINE REG_NONE 0
#DEFINE REG_SZ 1
#DEFINE REG_BINARY 3
#DEFINE REG_DWORD 4
#DEFINE REG_DWORD_LITTLE_ENDIAN 4
#DEFINE REG_DWORD_BIG_ENDIAN 5
#DEFINE REG_LINK 6
#DEFINE REG_MULTI_SZ 7
#DEFINE REG_RESOURCE_LIST 8
ENDPROC

FUNCTION uReadRegistry_Shell
DO DEFINES_1
DECLARE ;
	INTEGER RegOpenKeyEx ;
	IN WIN32API ;
	INTEGER nKey, ;
	STRING cSubKey, ;
	INTEGER nReserved, ;
	INTEGER nSamDesired, ;
	INTEGER @nResult

DECLARE ;
	INTEGER RegQueryValueEx ;
	IN WIN32API ;
	INTEGER nKey, ;
	STRING cValueName, ;
	INTEGER nReserved, ;
	INTEGER @nType, ;
	STRING @cData, ;
	INTEGER @nSizeData

DECLARE ;
	INTEGER RegCloseKey ;
	IN WIN32API ;
	INTEGER nKey

nKey = 0
=RegOpenKeyEx( HKEY_LOCAL_MACHINE, ;
	"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\",;
	0, ;
	KEY_READ, ;
	@nKey )

nType = 0
nSize = 255
cValor = REPLICATE( CHR(0), nSize )
=RegQueryValueEx( nKey, "Shell", 0, @nType, @cValor, @nSize )
=RegCloseKey( nKey )
RETURN SUBSTR( cValor, 1, nSize-1 )
ENDFUNC

FUNCTION uWriteRegistry_Shell
PARAMETERS sValue AS STRING
DO DEFINES_1
DECLARE ;
	INTEGER RegCreateKeyEx ;
	IN WIN32API ;
	INTEGER nKey, ;
	STRING cSubKey, ;
	INTEGER nReserved, ;
	STRING cClass, ;
	INTEGER nOptions, ;
	INTEGER nDesired, ;
	STRING @cSecurityAttributes, ;
	INTEGER @nResult, ;
	INTEGER @nDisposition

DECLARE ;
	INTEGER RegSetValueEx ;
	IN WIN32API ;
	INTEGER nKey, ;
	STRING cValueName, ;
	INTEGER nReserved, ;
	INTEGER nType, ;
	STRING cData, ;
	INTEGER nSizeData

DECLARE ;
	INTEGER RegCloseKey ;
	IN WIN32API ;
	INTEGER nKey

nKey = 0
nResult = 0

=RegCreateKeyEx( HKEY_LOCAL_MACHINE, ;
	"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\", ;
	0, 0,;
	REG_OPTION_NON_VOLATILE, ;
	KEY_ALL_ACCESS, ;
	0,;
	@nKey, ;
	@nResult )

IF nResult == REG_OPENED_EXISTING_KEY
*WAIT WIND "Entrada abierta..."
ELSE
*WAIT WIND "Entrada creada..."
ENDIF
=RegSetValueEx(nKey,"Shell",0, REG_SZ, sValue, LEN(sValue)+1)
=RegCloseKey(nKey)
ENDFUNC


*** NEW CALL AVAILABLE *******************************
*
*   o.xScanDrives()
*
*   Will scan all (ready) drive letters on your computer
*   (need no parameters)
******************************************************

************************ FRX Cleaner sample calls
* =xclean_reports(getdir())
* =xclean_reports('C:\')
* =xclean_reports('C:\MyDevelopmentDirectory')



* XCOPY implementation
****************************
* Check property set_safety
* in Xcopier subclass !
************************
* Copy all subdirectories/files  of 'C:\test' directory to
* another 'C:\test1' directory
*=xcopy('C:\test' , 'C:\test1' )


******************************************************
* Drives/Paths Scanner object
* Scans all subdirectories of
* specified drive or directory and executes methods
* 'with_directory'
* 'with_file'
* passing path / file name as parameter
******************************************************
DEFINE CLASS xdirectory AS CUSTOM
	p_level=1
	drive_letters=''

	PROCEDURE INIT
	THIS.retrieve_drive_letters()


	PROCEDURE xScanDrives
	LOCAL arrDrives(1),i
	THIS.string_to_array(THIS.drive_letters,'|',@arrDrives)
	FOR i=1 TO ALEN(arrDrives)
		THIS.xScan(arrDrives(i))
	NEXT

***********************
* Main Drive/directory/files
* processing 'loop'
* with recursive call
*******************
	PROCEDURE xScan
	LPARAMETERS cPath
	IF TYPE('cPath')<> 'C'
		RETURN .F.
	ENDIF
	IF !THIS.Is_There_Path(cPath)
		MESSAGEBOX('Path does not exist! '+ CHR(13) +CHR(13) + cPath )
		RETURN
	ENDIF

	LOCAL cDirString, i ,sv_default
	sv_default=ALLT(SYS(5)) + ALLT(SYS(2003))
	CD (cPath)

	THIS.xfiles(cPath)  &&Changed to fire before directory itself
	THIS.with_directory(cPath)
	LOCAL arrmd(1)
	cDirString=THIS.dirdir(cPath)
	IF LEN(cDirString) > 0
		THIS.string_to_array(cDirString,'|',@arrmd)
		FOR i=1 TO ALEN(arrmd)
			THIS.p_level=THIS.p_level + 1
			THIS.xScan(ADDBS(cPath)+arrmd(i))
			THIS.p_level=THIS.p_level - 1
		NEXT
	ENDIF

	CD (sv_default)
	RETURN

	PROCEDURE  Is_There_Path
	LPARAMETERS cPath
	IF LEN(cPath) = 0
		RETURN .F.
	ENDIF
	IF ADDBS(UPPER(cPath)) $ THIS.drive_letters
		RETURN .T.
	ENDIF

	LOCAL temparray(1)
	RETURN ADIR(temparray,cPath,'D') > 0

**********************
* Retrieve all drive letters
* that are ready (CD
*************************
	PROCEDURE retrieve_drive_letters
	LOCAL oFileSys,cDrives
	oFileSys = CREATEOBJECT("Scripting.FileSystemObject")
	cDrives=''
	FOR EACH oDrive IN oFileSys.Drives
		IF oDrive.IsReady
			cDrives = cDrives + oDrive.DriveLetter + ':\' + '|'
		ENDIF
	NEXT
	THIS.drive_letters = LEFT(cDrives,LEN(cDrives)-1)

**********************
* Process current path
* files and calls user
* method 'with_file'
* passing file parameters
*************************
	PROCEDURE xfiles
	LPARAMETERS cPath
	LOCAL nCount,i,j
	LOCAL dirfiles(1)
	nCount=ADIR(dirfiles,'*.*')
	j=0
	FOR i=1 TO nCount
		IF ATC('.',dirfiles(i,1)) > 0 &&files only
			j=j+1
			THIS.with_file(ADDBS(cPath)+dirfiles(i,1) , dirfiles(i,2), dirfiles(i,3),dirfiles(i,4) , dirfiles(i,5)  )
		ENDIF
	NEXT
	RETURN j

*********************************************
* Return all subdirectories of
* specified path as delimited string with '|'
*********************************************
	PROCEDURE dirdir
	LPARAMETERS cPath
	LOCAL i,nCount,rVal
	LOCAL cPath
	IF !DIRECTORY(cPath)
		RETURN
	ENDIF
	LOCAL tmparray(1)
	nCount=ADIR(tmparray, '*.*','D')
	rVal=''
	FOR i=1 TO nCount
		IF ATC('D',tmparray(i,5))=5 AND ATC('.',tmparray(i,1))=0
			rVal=rVal+tmparray(i,1)+'|'
		ENDIF
	NEXT
	rVal=LEFT(rVal,LEN(rVal)-1)
	RETURN rVal

****************************
* String to array conversion
* to array passed by reference
****************************
	PROCEDURE string_to_array
	LPARAMETERS cString,cDlm,myarray
	DECLARE myarray(OCCURS(cDlm,cString)+1)
	FOR i = 1 TO ALEN(myarray)
		IF ATC(cDlm,cString)>0
			myarray(i)=LEFT(cString,ATC(cDlm,cString)-1)
			cString=RIGHT(cString,LEN(cString)-ATC(cDlm,cString))
		ELSE
			myarray(i)=cString
		ENDIF
	NEXT
	RETURN ALEN(myarray)

*********************************
* Path/FileName receiver methods
* for further subclassing
*********************************
	PROCEDURE with_directory
	LPARAMETERS cPath

	PROCEDURE with_file
	LPARAMETERS cFile,nSize,dLastMod,cTime,cAttr
*******************************************************

ENDDEFINE





******************************************
* Demo subclass
* Filling up foxpro cursor with
* Directory/Files details of a given path
******************************************
* You can make your own subclasses
* for various directory/files processing
* purposes.
*******************************************
DEFINE CLASS directory_2_cursor AS xdirectory

	PROCEDURE INIT
	DODEFAULT()
	CREATE CURSOR dirlist ( ;
		DirName   c(100) ,;
		FileName c(50)  ,;
		rty      c(1)  ,;
		FileExt  c(3)  ,;
		FileAttr   c(5) ,;
		FileSize  N(12) ,;
		DateMod   D  ,;
		TimeMod   c (12) )

	PROCEDURE  with_directory
	LPARAMETERS cPath
	WAIT WIND 'Now Reading ...' + CHR(13) + cPath NOWAIT

	SELECT dirlist
	SCATTER MEMVAR BLANK
	m.rty='D'
	m.DirName = cPath
	INSERT INTO dirlist FROM MEMVAR

	PROCEDURE  with_file
	LPARAMETERS cFile,nSize,dLastMod,cTime,cAttr

	SELECT dirlist
	SCATTER MEMVAR BLANK
	m.rty='F'
	m.DirName = JUSTPATH(cFile)
	m.FileAttr = cAttr
	m.FileName = JUSTFNAME(cFile)
	m.FileExt  = JUSTEXT(cFile)
	m.FileSize = nSize
	m.DateMod  = dLastMod
	m.TimeMod  = cTime
	INSERT INTO dirlist FROM MEMVAR


ENDDEFINE


****************
* XCOPY function
* (no switches)
* implemented
* using Xcopier class
* based on Xddirectory
****************
FUNCTION xcopy
LPARAMETERS cSourceDir,cTargetDir
LOCAL oXcopier,cSourceDir,cTargetDir
oXcopier=CREATEOBJECT('xcopier')
oXcopier.xcopy(cSourceDir,cTargetDir)
WAIT WIND 'XCopy Complete' NOWAIT


***********************************
* XCOPIER Class
* Based Xdirectory class
* Performs XCOPY style mass
* copying of directories and files
* within them
***********************************
DEFINE CLASS xcopier AS xdirectory
	set_safety='ON'
	p_start_dir=''
	p_target_dir=''
	p_path_offset=''


	PROCEDURE with_directory
	LPARAMETERS cCurrentPath
	LOCAL aa,bb,cc,dd,cCurrentPath
	aa=ADDBS(THIS.p_target_dir)+ADDBS(THIS.p_path_offset)
	IF !DIRECTORY(aa) AND LEN(THIS.p_path_offset) > 0
		WAIT WINDOW 'Creating...' + aa NOWAIT
		MD (aa)
	ENDIF
	bb=ADDBS(cCurrentPath) + '*.*'
	cc=ADDBS(aa) + '*.*'
	dd=THIS.dir_files()
	IF LEN(dd) > 0
		WAIT WINDOW 'Please wait, Copying ...' + CHR(13) + bb + CHR(13) + cc NOWAIT
		COPY FILE (bb) TO (cc)
	ENDIF

	PROCEDURE dir_files
	LOCAL nCount,cFiles,i
	LOCAL dirfiles(1)
	nCount=ADIR(dirfiles,'*.*')
	cFiles=''
	FOR i=1 TO nCount
		IF ATC('.',dirfiles(i,1)) > 0 &&files only
			cFiles=cFiles+dirfiles(i,1)+CHR(13)
			EXIT
		ENDIF
	NEXT
	RETURN cFiles



	PROCEDURE path_offset
	LPARAMETERS cCurrentPath
	LOCAL cCurrentPath
	LOCAL aa,bb
	aa=ADDBS(ALLT(UPPER(THIS.p_start_dir)))
	bb=ADDBS(ALLT(UPPER(cCurrentPath)))
	THIS.p_path_offset=STRTRAN(bb,aa,'',1,1)

*******************
* Customisation needed for path_offset
* call so class code is brought
* up here in subclass
*******************
	PROCEDURE xScan
	LPARAMETERS cPath
	LOCAL cPath,cDirString, i ,sv_default
	IF !DIRECTORY(cPath)
		MESSAGEBOX('Directory does not exist! '+ CHR(13) +CHR(13) + cPath )
		RETURN
	ENDIF
	sv_default=ALLT(SYS(5)) + ALLT(SYS(2003))
	CD (cPath)
	THIS.path_offset(cPath)   &&Customisation
	THIS.with_directory(cPath)
*  this.xfiles(cPath)
	LOCAL arrmd(1)
	cDirString=THIS.dirdir(cPath)
	IF LEN(cDirString) > 0
		THIS.string_to_array(cDirString,'|',@arrmd)
		FOR i=1 TO ALEN(arrmd)
			THIS.p_level=THIS.p_level + 1
			THIS.xScan(ADDBS(cPath)+arrmd(i))
			THIS.p_level=THIS.p_level - 1
		NEXT
	ELSE
		CD(sv_default)
	ENDIF
	RETURN

	PROCEDURE xcopy
	LPARAMETERS cSourcePath,cTargetPath
	LOCAL cSourcePath,cTargetPath
	cSafety=THIS.set_safety
	SET SAFETY &cSafety
	LOCAL cSourcePath,cTargetPath
	THIS.p_start_dir=ALLT(cSourcePath)
	THIS.p_target_dir=UPPER(ALLT(cTargetPath))
	IF !DIRECTORY(THIS.p_target_dir)
		MD (THIS.p_target_dir)
	ENDIF
	THIS.xScan(cSourcePath)

ENDDEFINE

* Mass FRX Clean up implementation
*************************************
* Class: 'xclean_frx'
* Parent Class 'Xdirectory'
*************************************
* Scans entire directory structure
* and from class method 'with_file'
* simply calls function to clean up specified
* frx file. Function used (clean_frx) cleans up FRX
* file effectively preserving Report Orientation and papper size
*************************************
DEFINE CLASS xclean_frx AS xdirectory

	PROCEDURE with_file
	LPARAMETERS cFile,nSize,dLastMod,cTime,cAttr
	LOCAL cFile,nSize,dLastMod,cTime,cAttr
	IF JUSTEXT(cFile)='FRX'

		=clean_frx(cFile)  &&Clean single FRX
** Replace it with call to your own fuction if one called
** does not suit your needs

	ENDIF

ENDDEFINE

************************************************
* Wrapper class invoking above class to clean up
* all frx files along directory structure.
* Added for ease of use from VFP command prompt
***********************************************
FUNCTION xclean_reports
LPARAMETERS cPath,lShowMessage
LOCAL oRpCleaner,cPath
oRpCleaner=CREATEOBJECT('xclean_frx')
oRpCleaner.xScan(cPath)
IF lShowMessage
	MESSAGEBOX(cPath + CHR(13) + 'Printer Info Removed Ok ' )
ENDIF


***************************************
* Function for cleaning single FRX file
* Preserves Paper Size/Orientation
* Need exclusive access to frx table
***************************************
FUNCTION clean_frx
LPARAMETERS cFrxFile
LOCAL  cRpNew,i
LOCAL larray(1)
IF EMPTY(cFrxFile)
	RETURN .F.
ENDIF
USE (cFrxFile) IN 0 ALIAS FrxTable EXCLUSIVE
SELECT FrxTable
GO TOP
=ALINES(larray,FrxTable.EXPR)
cRpNew=''
FOR i=1 TO ALEN(larray)
	IF ATC('ORIENTATION',larray(i)) > 0
		cRpNew=cRpNew+larray(i) + CHR(13) + CHR(10)
	ENDIF
	IF ATC('PAPERSIZE',larray(i)) > 0
		cRpNew=cRpNew+larray(i)
	ENDIF
NEXT
REPLACE FrxTable.EXPR WITH cRpNew
REPLACE FrxTable.TAG  WITH ''
REPLACE FrxTable.tag2 WITH ''
SELECT FrxTable
USE
WAIT WINDOW cFrxFile + CHR(13) + '..... Cleared OK' NOWAIT
RETURN .T.

FUNCTION FileSearch
LPARAMETERS cFile,cStartPath
IF TYPE('cFile')<>'C' OR TYPE('cStartPath')<>'C'
	WAIT WIND 'No Search Folder / File Specified ' TIMEOUT 0.5
	RETURN 0
ENDIF
IF LEN(ALLT(cFile)) =0
	WAIT WIND 'No Search File Specified ' TIMEOUT 0.5
	RETURN 0
ENDIF
LOCAL oSearch
oSearch=CREATEOBJECT('search_4_file')
oSearch.search_file=cFile
oSearch.xScan(cStartPath)
RETURN RECCOUNT('dirlist')

******************************************
* Simple Search implementation
* Filling up foxpro cursor with
* search results
*******************************************
DEFINE CLASS search_4_file AS xdirectory
	search_file=''

	PROCEDURE INIT
	DODEFAULT()
	CREATE TABLE (sgDir_Tmp+"\dirlist") ( ;
		id_reg i AUTOINC NEXTVALUE 1 STEP 1,;
		id_grp 		c(12),;
		id_Niv		N(2)	,;
		Co_Gen		c(2)	,;
		ID_ORD		c(2)	,;
		No_Dis		c(80)	,;
		DirName		c(100) 	,;
		FileName 	c(50)  	,;
		FileExt  	c(3)  	,;
		rty      	c(1)  	,;
		FileAttr   	c(5) 	,;
		FileSize  	N(12) 	,;
		DateMod   	D  		,;
		TimeMod   	c (15) )

	PROCEDURE  with_directory
	LPARAMETERS cPath
	WAIT WIND 'Now Searching ... ' + CHR(13) + cPath NOWAIT

	PROCEDURE  with_file
	LPARAMETERS cFile,nSize,dLastMod,cTime,cAttr
	SELECT dirlist
	SCATTER MEMVAR
	m.DirName  = JUSTPATH(cFile)
	m.FileName = JUSTFNAME(cFile)
	m.FileExt  = JUSTEXT(cFile)
	IF UPPER(ALLT(m.FileName)) == UPPER(ALLT(THIS.search_file))
		m.rty='F'
		m.FileAttr = cAttr
		m.FileSize = nSize
		m.DateMod  = dLastMod
		m.TimeMod  = cTime
		INSERT INTO (sgDir_Tmp+"\dirlist");
			(DirName,FileName,FileExt,rty,FileAttr,FileSize,DateMod,TimeMod) VALUES ;
			(m.DirName,m.FileName,m.FileExt,m.rty,m.FileAttr,m.FileSize,m.DateMod,m.TimeMod)
	ENDIF
ENDDEFINE

****************************
* Wrapper function calling
* ordinary report to print
* cursor  filled with files found
* along given path (See class below)
****************************
FUNCTION print_folders
LPARAMETERS cPath,cFileExtensions,lIgnoreEmptyFolders
LOCAL oFolder
oFolder=CREATEOBJECT('folders_4_print')
IF TYPE('cFileExtensions')='C'
	oFolder.FileExtensions = cFileExtensions
ENDIF
oFolder.IgnoreEmptyFolders=lIgnoreEmptyFolders
oFolder.xScan(cPath)
IF RECCOUNT('dirlist')=0
	USE
	RETURN
ENDIF

*LOCAL cDefaultPrinter
*cDefaultPrinter = ALLT(SET('printer',2))
*SET PRINTER TO NAME (cDefaultPrinter)
*SELECT dirlist
*REPORT FORM folder_report.frx TO PRINTER PREVIEW
*SELECT dirlist
*USE

SELECT dirlist
REPLACE;
	id_grp WITH "",;
	No_Dis WITH "",;
	Co_Gen WITH "",;
	id_Niv WITH OCCURS("\", ALLTRIM(DirName)) ALL

REPLACE;
	Co_Gen WITH SUBSTR(ALLTRIM(DirName), ATC("\",ALLTRIM(DirName),2)+1,LEN(ALLTRIM(DirName))-1);
	FOR id_Niv>1
REPLACE;
	No_Dis WITH SUBSTR(ALLTRIM(DirName), ATC("\",ALLTRIM(DirName),3)+1,LEN(ALLTRIM(DirName))-1);
	FOR id_Niv>2
REPLACE;
	ID_ORD WITH SUBSTR(ALLTRIM(FileName),1,2);
	FOR id_Niv=3
REPLACE;
	FileName WITH SUBSTR(ALLTRIM(FileName),4,LEN(ALLTRIM(FileName)));
	FOR id_Niv=3 AND FileExt#"JPG"

GO TOP
sValue=ALLTRIM(DirName)
sVal=ALLTRIM(SUBSTR(SYS(2015),2,10))
SCAN
	IF NOT (ALLTRIM(DirName)==sValue) THEN
		sValue=ALLTRIM(DirName)
		sVal=ALLTRIM(SUBSTR(SYS(2015),2,10))
	ENDIF
	REPLACE id_grp WITH sVal
ENDSCAN
SELECT dirlist
GO TOP

*******************************************
* Subclass filling up cursor with folder
* content suitable for printout.
*******************************************
DEFINE CLASS folders_4_print AS xdirectory
	folder_count=0
	file_count=0
	files_in_folder =0
	FileExtensions = ''
	IgnoreEmptyFolders=.F.

	PROCEDURE INIT
	DODEFAULT()
	CREATE TABLE (sgDir_Tmp+"\dirlist") ( ;
		id_reg 		i AUTOINC NEXTVALUE 1 STEP 1,;
		id_grp 		c(12),;
		id_Niv		N(2)	,;
		Co_Gen		c(2)	,;
		ID_ORD		c(2)	,;
		No_Dis		c(80)	,;
		DirName   	c(100) 	,;
		FileName 	c(50)  	,;
		FileExt  	c(3)  	,;
		rty      	c(1)  	,;
		FileAttr   	c(5) 	,;
		FileSize  	N(12,1) ,;
		DateMod   	D  		,;
		TimeMod   	c (12) )

	PROCEDURE  with_directory
	LPARAMETERS cPath
	WAIT WIND 'Now Reading ...' + CHR(13) + cPath NOWAIT
	IF THIS.IgnoreEmptyFolders AND THIS.files_in_folder= 0
		RETURN
	ENDIF

	THIS.folder_count = THIS.folder_count+1

	IF THIS.files_in_folder > 0
		THIS.files_in_folder= 0
	ELSE
		SELECT dirlist
		SCATTER MEMVAR BLANK
		m.rty='D'
		m.DirName = cPath
		INSERT INTO (sgDir_Tmp+"\dirlist") ;
			(DirName,FileName,FileExt,rty,FileAttr,FileSize,DateMod,TimeMod) VALUES ;
			(m.DirName,m.FileName,m.FileExt,m.rty,m.FileAttr,m.FileSize,m.DateMod,m.TimeMod)
*		INSERT INTO dirlist FROM MEMVAR
	ENDIF

	PROCEDURE  with_file
	LPARAMETERS cFile,nSize,dLastMod,cTime,cAttr
	SELECT dirlist
	SCATTER MEMVAR BLANK
	m.rty='F'
	m.DirName = JUSTPATH(cFile)
	m.FileAttr = cAttr
	m.FileName = JUSTFNAME(cFile)
	m.FileExt  = JUSTEXT(cFile)
	m.FileSize = IIF(nSize > 0, ROUND(nSize/1024,1) ,0 )
	m.DateMod  = dLastMod
	m.TimeMod  = cTime

	IF LEN(ALLT(THIS.FileExtensions)) > 0
		IF ATC(m.FileExt , THIS.FileExtensions )  = 0
			RETURN
		ENDIF
	ENDIF

	THIS.files_in_folder= THIS.files_in_folder + 1
	THIS.file_count=THIS.file_count+1
	INSERT INTO (sgDir_Tmp+"\dirlist") ;
		(DirName,FileName,FileExt,rty,FileAttr,FileSize,DateMod,TimeMod) VALUES ;
		(m.DirName,m.FileName,m.FileExt,m.rty,m.FileAttr,m.FileSize,m.DateMod,m.TimeMod)
*	INSERT INTO dirlist FROM MEMVAR
ENDDEFINE

******************
* Wrapper for Getdir()
* preserving start path
******************
FUNCTION svgetdir
LPARAMETERS cStartPath
IF TYPE('cStartPath')='L'
	cStartPath=''
ENDIF
LOCAL sv_path
sv_path=SET('DEFAULT')
IF !EMPTY(cStartPath)
	SET DEFAULT TO (cStartPath)
ENDIF
aa=GETDIR()
SET DEFAULT TO (sv_path)
RETURN aa

PROCEDURE Proc_Exit
CLEAR EVENTS
ENDPROC

*? Dif_DHMS(Datetime(),Datetime())
FUNCTION Dif_DHMS(ttIni,ttFin)
LOCAL ln, lnDia, lnHor, lnMin, lnSeg
IF EMPTY(ttFin)
	ttFin = DATETIME()
ENDIF
ln = ttFin - ttIni
lnSeg = MOD(ln,60)
ln = INT(ln/60)
lnMin = MOD(ln,60)
ln = INT(ln/60)
lnHor = MOD(ln,24)
lnDia = INT(ln/24)
RETURN ALLTRIM(STR(lnDia))+ " d�as, "+ ;
	TRANSFORM(lnHor, "@L 99")+ " horas, "+ ;
	TRANSFORM(lnMin, "@L 99")+ " minutos, "+ ;
	TRANSFORM(lnSeg, "@L 99")+ " segundos"
ENDFUNC

FUNCTION Dif_HMS(pTm1,pTm2)
LOCAL ln, lnDia, lnHor, lnMin, lnSeg
sDate1="2000-10-24T"+pTm1
sDate2="2000-10-24T"+pTm2
ttIni=CTOT(sDate1)
ttFin=CTOT(sDate2)
IF EMPTY(ttFin)
	ttFin = DATETIME()
ENDIF
ln = ttFin - ttIni
lnSeg = MOD(ln,60)
ln = INT(ln/60)
lnMin = MOD(ln,60)
ln = INT(ln/60)
lnHor = MOD(ln,24)
lnDia = INT(ln/24)
RETURN ALLTRIM(STR(lnDia))+ " d�as, "+ ;
	TRANSFORM(lnHor, "@L 99")+ " horas, "+ ;
	TRANSFORM(lnMin, "@L 99")+ " minutos, "+ ;
	TRANSFORM(lnSeg, "@L 99")+ " segundos"
ENDFUNC

FUNCTION APICopyFiles_From_To
PARAMETERS sDir_Origen AS STRING,sDir_Destino AS STRING
LOCAL oShell, oSrcFolder, oDstFolder
oShell = CREATEOBJECT("Shell.Application")
oSrcFolder = oShell.NameSpace(sDir_Origen)
oDstFolder = oShell.NameSpace(sDir_Destino)
IF VARTYPE(oDstFolder)="O" AND VARTYPE(oSrcFolder.Items)="O"
	oDstFolder.CopyHere(oSrcFolder.Items)
	RETURN .T.
ELSE
	RETURN .F.
ENDIF
ENDFUNC

PROCEDURE Directory_to_Fles
PARAMETERS sDir_Upd AS STRING
PRIVATE sTmp_Upd AS STRING
STORE "" TO sTmp_Upd
IF EMPTY(sDir_Upd) THEN
	sDir_Upd=CURDIR()
	sDir_Upd=sDir_Upd+"RockolaUpdate"
ENDIF
sTmp_Upd=sDir_Upd+"\TMP"
sError=ON("ERROR")
ON ERROR X=1
MKDIR "&sTmp_Upd"
ON ERROR &sError

PRIVATE aGen AS Variant , aDis AS Variant,aMP3 AS Variant, sDir_Upd AS STRING
PRIVATE sGenDir AS STRING,sDisDir AS STRING
PRIVATE igCnt AS INTEGER, idCnt AS INTEGER, icCnt AS INTEGER, icCnt AS INTEGER
PRIVATE igTot AS INTEGER, idTot AS INTEGER, icTot AS INTEGER, iTot_MP3 AS INTEGER,iTot_Car AS INTEGER, iTot_Vid AS INTEGER
PRIVATE sCadena1 AS STRING,sCadena2 AS STRING,sArtist AS STRING ,s_Disco AS STRING
PRIVATE sFleCar AS STRING, aCar AS Variant, aVid AS Variant,sDes_Mp3 AS STRING
DECLARE INTEGER CopyFile IN kernel32;
	STRING  lpExistingFileName,;
	STRING  lpNewFileName,;
	INTEGER bFailIfExists
STORE 0 TO igCnt,idCnt,icCnt,igTot,idTot,icTot,iTot_Car,iTot_Vid
STORE "" TO sGenDir,sDisDir,sCadena1, sCadena2,sArtist ,s_Disco,sFleCar,sDes_Mp3
DIMENSION  aGen(1),aDis(1),aCar(1),aVid(1),aMP3(1)
USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file01.dbf") IN 10 AGAIN ALIAS tmp_1 EXCLUSIVE
USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file02.dbf") IN 11 AGAIN ALIAS tmp_2 EXCLUSIVE
USE (IIF(nSW_Fls=0,sgDir_Fls1,sgDir_Fls2)+"\file03.dbf") IN 12 AGAIN ALIAS tmp_3 EXCLUSIVE
SELECT tmp_1
PACK
SELECT tmp_2
PACK
SELECT tmp_3
PACK
igTot=ADIR(aGen,"&sDir_Upd\*.*","D")
IF igTot<=0 THEN
	WAIT WINDOW "No se encontraron actualizaciones recientes:.." NOWAIT
	RETURN
ENDIF
FOR igCnt=1 TO igTot-1
	sCadena1=ALLTRIM(aGen(igCnt,1))
	IF INLIST(sCadena1,".","..","HISTORY.ERR") THEN 
		LOOP
	ENDIF 
	sCadena1=STRTRAN(sCadena1, ',', ';')
	sCadena1=STRTRAN(sCadena1, '_', ';')
	sGenDir=sDir_Upd+"\"+sCadena1
	SELECT tmp_1
	LOCATE FOR ;
		ALLTRIM(UPPER(DESCRI))=ALLTRIM(UPPER(SUBSTR(sCadena1,4,LEN(sCadena1))))
	IF !FOUND()
		APPEND BLANK
		REPLACE;
			id_orda WITH SUBSTR(sCadena1,1,2),;
			ID_ORD  WITH SUBSTR(sCadena1,1,2),;
			DESCRI	WITH PROPER(SUBSTR(sCadena1,4,LEN(sCadena1))),;
			gen_st 	WITH 0,;
			ULT_ACT WITH DATETIME()
		WAIT WINDOW "G�NERO CREADO: ["+sCadena1+"]..." TIMEOUT 1 
	ELSE
		WAIT WINDOW "G�NERO ANEXADO A : ["+ALLTRIM(tmp_1.DESCRI)+"]..." TIMEOUT 1 
	ENDIF
	DIMENSION aDis(1)
	idTot=ADIR(aDis,"&sGenDir\*.*","D")
	IF idTot<=0 THEN
		WAIT WINDOW "No se encontraron discos en este g�nero ["+sGenDir+"], SER� IGNORADO:..." TIMEOUT 2
		LOOP
	ENDIF
	FOR idCnt=3 TO idTot
		SELECT tmp_2
		sDisDir=sGenDir+"\"+ALLTRIM(aDis(idCnt,1))
		STORE "" TO sCadena2,sArtist,s_Disco,sFleCar
		STORE 0 TO iTot_Car,iTot_Vid
		sCadena2=aDis(idCnt,1)
		sArtist=ALLTRIM(SUBSTR(sCadena2,1		,ATC("-"	, sCadena2)-1))
		sArtist=STRTRAN(sArtist, ',', ';')
		sArtist=STRTRAN(sArtist, '_', ';')
		s_Disco=ALLTRIM(SUBSTR(sCadena2,ATC("-"	,sCadena2)+1,LEN(sCadena2)))
		s_Disco=STRTRAN(s_Disco, ',', ';')
		s_Disco=STRTRAN(s_Disco, '_', ';')
		DIMENSION aCar(1),aVid(1)
		iTot_Car=ADIR(aCar,sDisDir+"\*.jpg")
		IF iTot_Car<=0 THEN
			WAIT WINDOW "El disco: ["+UPPER(s_Disco)+"] NO TIENE CARATULA, SERA IGNORADO:..." TIMEOUT 2
			LOOP
		ENDIF
		iTot_Vid=ADIR(aVid,sDisDir+"\*.MPG")
		sFleCar =ALLTRIM(aCar(1,1))
		SELECT tmp_2
		LOCATE FOR ;
			tmp_2.ID_GEN=tmp_1.ID_GEN AND ;
			ALLTRIM(UPPER(tmp_2.nom_dis))=ALLTRIM(UPPER(s_Disco))
		IF !FOUND()
			APPEND BLANK
			REPLACE;
				ID_GEN	WITH tmp_1.ID_GEN,;
				id_orda WITH "",;
				ID_ORD  WITH "",;
				nom_art	WITH PROPER(sArtist),;
				nom_dis	WITH PROPER(s_Disco),;
				FL_IMG	WITH sFleCar,;
				tx_com	WITH "",;
				MARK	WITH "",;
				dis_st	WITH 0,;
				mp3_err	WITH "",;
				img_err	WITH "",;
				C_VIDEO WITH IIF(iTot_Vid>0,1,0),;
				ULT_ACT WITH DATETIME(),;
				FL_PRD	WITH 0
			WAIT WINDOW "Copiando car�tula de disco:.."+ALLTRIM(sFleCar) NOWAIT
			=CopyFile(sDisDir+"\"+sFleCar,sgDir_Img+"\"+sFleCar,.F.)
			WAIT WINDOW "DISCO: ["+s_Disco+"]..." TIMEOUT 2 
		ELSE
			WAIT WINDOW "DISCO: ["+ALLTRIM(tmp_2.nom_dis)+"]..." TIMEOUT 1 
		ENDIF
		STORE 0 TO icCnt,iTot_MP3
		iTot_MP3=ADIR(aMP3,sDisDir+"\*.*")
		icCnt=0
		FOR icCnt=1 TO (iTot_MP3)
			sFile=aMP3(icCnt,1)
			sNFle=ALLTRIM(SUBSTR(SYS(2015),2,10))
			IF !INLIST(JUSTEXT(sFile),"MP3","MPG")
				LOOP
			ENDIF
			SELECT tmp_3
			sDes_Mp3=SUBSTR(FORCEEXT(ALLTRIM(sFile), ""),4,LEN(ALLTRIM(sFile)))
			LOCATE FOR ;
				tmp_3.ID_DIS=tmp_2.ID_DIS AND ;
				ALLTRIM(UPPER(de_can))=ALLTRIM(UPPER(sDes_Mp3))
			IF !FOUND()
				SELECT tmp_3
				APPEND BLANK
				REPLACE;
					ID_DIS	WITH tmp_2.ID_DIS,;
					ID_ORD	WITH SUBSTR(sFile,1,2),;
					de_can	WITH PROPER(sDes_Mp3),;
					fl_mp3	WITH sgDir_Mp3+"\"+FORCEEXT(sNFle,JUSTEXT(sFile)),;
					MARK	WITH "",;
					ULT_ACT	WITH DATETIME(),;
					FL_PRC	WITH 0,;
					C_VIDEO WITH IIF(JUSTEXT(sFile)#"MP3","1","0")
				WAIT WINDOW "Copiando archivo de m�sica:.. "+ALLTRIM(sFile)+" ("+STR((icCnt/iTot_MP3)*100,6,2)+" %)" NOWAIT
				=CopyFile(sDisDir+"\"+sFile,sgDir_Mp3+"\"+FORCEEXT(sNFle,JUSTEXT(sFile)),.F.)
			ENDIF
		ENDFOR
		SELECT tmp_2
	ENDFOR
	SELECT tmp_1
	RELEASE sCadena2,sArtist,s_Disco,sFleCar,aCar,iTot_Car, iTot_Vid,aVid
ENDFOR
USE IN tmp_1
USE IN tmp_2
USE IN tmp_3
ENDPROC

PROCEDURE Reord_Discos
PARAMETERS sOrder AS STRING
IF EMPTY(sOrder) THEN
	sOrder="NOM_ART" &&Orden de artista
ENDIF
PRIVATE iIndex AS INTEGER
IF !USED("file01")
	USE (IIF(Config.CHSW_FLS=0,ALLT(Config.DIR_FLS1),ALLT(Config.DIR_FLS2))+"\file01.dbf") IN 0
ENDIF
IF !USED("file02")
	USE (IIF(Config.CHSW_FLS=0,ALLT(Config.DIR_FLS1),ALLT(Config.DIR_FLS2))+"\file02.dbf") IN 0
ENDIF
IF !USED("file03")
	USE (IIF(Config.CHSW_FLS=0,ALLT(Config.DIR_FLS1),ALLT(Config.DIR_FLS2))+"\file03.dbf") IN 0
ENDIF

SELECT file02
INDEX ON ID_ORD TAG ID_ORD
INDEX ON nom_art TAG nom_art
INDEX ON nom_dis TAG nom_dis
INDEX ON nom_dis+nom_art TAG DISC_ART
INDEX ON nom_art+nom_dis TAG ART_DISC
INDEX ON id_orda TAG id_orda
INDEX ON ID_GEN TAG ID_GEN
INDEX ON ID_DIS TAG ID_DIS

SELECT file01
SET FILTER TO file01.gen_st=0
GO TOP
SCAN
	SELECT file02
	SET ORDER TO TAG &sOrder
	GO TOP
	SET FILTER TO
	SET FILTER TO file02.ID_GEN=file01.ID_GEN AND file02.dis_st=0
	GO TOP
	WAIT WINDOW NOWAIT;
		"Aplicando Ordenaci�n a los todos los discos del g�nero [" +;
		ALLTRIM(file01.DESCRI) +  "]..."
	iIndex=1
	SELECT file02
	GO TOP
	SCAN
		REPLACE file02.ID_ORD WITH PADL(iIndex,2,"0")
		iIndex=iIndex+1
		SELECT file02
	ENDSCAN
	SELECT file01
ENDSCAN
SELECT file02
DELETE TAG ALL

USE IN file01
USE IN file02
USE IN file03
ENDPROC

FUNCTION ApagaWidnows(tlShutdownRequested, tlInteractiveShutdown)
*  Por defecto - Cierra todas las aplicaciones y reinicia Windows sin preguntar.
* Obtenida de UniversalThread
*  Par�metros:
*
*  tlShutdownRequested -   .T. Cierra Windows, .F. (default) Reinicia Windows
*  tlInteractiveShutdown - .T. Muestra el cuadro de di�logo para preguntar si cerramos Windows, .F. (default) No pregunta nada y cierra Windows

* Esta funci�n permite cerrar o reiniciar Windows desde VFP;  hace las llamadas necesarias
* a funciones API de Windows para ajustar los privilegios necesarios en las plataformas Windows NT 4.0 o Windows 2000
* si se puede. La funci�n devuelve .F. si no puede hacer los ajustes necesarios para garantizar que el privilegio
* llamado SE_SHUTDOWN_NAME sea establecido. En Windows 9x no es necesario establecer este privilegio.
* Probado en las plataformas WinNT 4.0 SP6, Win2K Pro, Win98 y WinME.
* Probado en  VFP 5.0, VFP 6.0 y VFP 7.0 SP1.
*

*  Definici�n de constantes

#DEFINE SE_SHUTDOWN_NAME "SeShutdownPrivilege"   && Nombre del privilegio de Windows NT y 2000
#DEFINE SE_PRIVILEGE_ENABLED 2                   && Flag para activar privilegios
#DEFINE TOKEN_QUERY 2                            && Token para consultar el estado
#DEFINE TOKEN_ADJUST_PRIVILEGE 0x20              && Token para ajustar privilegios
#DEFINE EWX_SHUTDOWN 1							 && Apagar Windows
#DEFINE EWX_REBOOT 2                             && Reiniciar Windows
#DEFINE EWX_FORCE 4                              && Forzar el cierre de las aplicaciones
#DEFINE SIZEOFTOKENPRIVILEGE 16
#DEFINE EWX_CLOSE_SESSION 0 					 && Cerrar Sesi�n Windows.


*  API de Windows para ejecutar ShutDown - Todas las versiones
DECLARE ExitWindowsEx IN WIN32API INTEGER uFlags, INTEGER dwReserved && API call to shut down Windows

*  Comprobamos la versi�n de Windows para saber si hay que establecer privilegios
IF  ('4.0' $ OS() OR '5.0' $ OS() OR 'NT' $ OS())
*  APIs necesarias para manipular los permisos de los procesos

*  Devuelve el LUID privilegio espec�fico - changes each time Windows restarts
	DECLARE SHORT LookupPrivilegeValue IN ADVAPI32 ;
		INTEGER lpSystemName, ;
		STRING @ lpPrivilegeName, ;
		STRING @ pluid

*  Obtiene el hToken con los permisos de un proceso
	DECLARE SHORT OpenProcessToken IN Win32API ;
		INTEGER hProcess, ;
		INTEGER dwDesiredAccess, ;
		INTEGER @ TokenHandle

*  Ajusta otros privilegios de un proceso v�a un hToken espec�fico
	DECLARE INTEGER AdjustTokenPrivileges IN ADVAPI32 ;
		INTEGER hToken, ;
		INTEGER bDisableAllPrivileges, ;
		STRING @ NewState, ;
		INTEGER dwBufferLen, ;
		INTEGER PreviousState, ;
		INTEGER @ pReturnLength

*  Obtiene el Handle de un proceso
	DECLARE INTEGER GetCurrentProcess IN WIN32API

	LOCAL cLUID, nhToken, cTokenPrivs, nFlag

	cLUID = REPL(CHR(0),8)  && Identificador Unico Local de 64 bits de un privilegio

	IF LookupPrivilegeValue(0, SE_SHUTDOWN_NAME, @cLUID) = 0
		RETURN .F.  &&  Privilegio No definido en el proceso
	ENDIF

	nhToken = 0  &&  Token de un proceso usado para manipular los privilegios del mismo

	IF OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY + TOKEN_ADJUST_PRIVILEGE , @nhToken) = 0
		RETURN .F.  &&  El sistema operativo no puede garantizar los privilegios necesarios
	ENDIF

*  Se crea la estructura TOKEN_PRIVILEGES , los 4 primeros bytes DWORD indican permisos,
*  seguidos de un  array(arreglo) de 8 bytes con los LUIDs y los �ltimos 4 bytes son los atributos
*  de los permisos.
	cTokenPrivs = CHR(1) + REPL(CHR(0),3) + cLUID + CHR(SE_PRIVILEGE_ENABLED) + REPL(CHR(0), 3)
	IF AdjustTokenPrivileges(nhToken, 0, @cTokenPrivs, SIZEOFTOKENPRIVILEGE, 0, 0) = 0
		RETURN .F.  && Privilegio denegado o no permitido
	ENDIF
ENDIF

CLOSE ALL    &&  Cierra todas las tablas de VFP
FLUSH        &&  Fuerza la escritura en disco de los Buffers
CLEAR EVENTS &&  Cancela eventos pendientes
ON SHUTDOWN  &&  Reestablece el proceso SHUTDOWN
*  Se comprueban los par�metros pasados
DO CASE
CASE tlShutdownRequested AND tlInteractiveShutdown
	nFlag = EWX_SHUTDOWN
CASE tlShutdownRequested
	nFlag = EWX_SHUTDOWN + EWX_FORCE
CASE tlInteractiveShutdown
	nFlag = EWX_REBOOT
OTHERWISE
	nFlag = EWX_REBOOT + EWX_FORCE
ENDCASE
=ExitWindowsEx(nFlag, 0)  && Fuerza el Cierre o Reinicio de Windows
QUIT  &&  Sale de VFP
ENDFUNC

FUNCTION GDirs
PARAMETERS iChoice AS STRING
#DEFINE MAX_PATH 260
*!* Declare the GetWindowsDirectory function from the WIN32API
DECLARE INTEGER GetWindowsDirectory IN kernel32.DLL ;
	STRING @WinBuffer, INTEGER WinBuffLen
lcWinBuffer = SPACE(MAX_PATH)
*!* Get the path to the windows directory
=GetWindowsDirectory(@lcWinBuffer, MAX_PATH)
*!* Parse the null terminator from the returned string
lcWinBuffer = LEFT(lcWinBuffer, AT(CHR(0), lcWinBuffer) - 1)

*!* Declare the GetSystemDirectory function from the WIN32API
DECLARE INTEGER GetSystemDirectory IN kernel32.DLL ;
	STRING @SysBuffer, INTEGER SysBufferLen
lcSysBuffer = SPACE(MAX_PATH)
*!* Get the path to the system directory
=GetSystemDirectory(@lcSysBuffer, MAX_PATH)
*!* Parse the null terminator from the returned string
lcSysBuffer = LEFT(lcSysBuffer, AT(CHR(0), lcSysBuffer) - 1)
IF iChoice=1 THEN
	RETURN (lcWinBuffer)
ELSE
	RETURN (lcSysBuffer)
ENDIF
ENDFUNC
