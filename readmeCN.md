# PBox ��һ������ Dll ��̬�ⴰ���ģ�黯����ƽ̨

- [English](readme.md)

## һ��������ּ
    ���ž������޸Ļ��޸�ԭ�й���Դ�����ԭ��;
    ֧�� Delphi��VC��QT Dll����;

## ��������ƽ̨
    Delphi10.3.3��WIN10X64 �¿�����
    ����û��ʹ���κε������ؼ���
    ������һЩ��Դ�⣬ͳһ������ 3rdparty Ŀ¼���棻ֱ�����ã����谲װ;
    WIN7X64��WIN10X64�²���ͨ����֧��X86��X64;
    ���䣺dbyoung@sina.com
    QQȺ��101611228

## ����ʹ�÷���
### Delphi��
* Delphi ԭ EXE �����ļ����޸�Ϊ Dll ���̡�����ض������Ϳ����ˣ�ԭ�д��벻�����κ��޸ġ�
* �ѱ����� Dll �ļ����õ� plugins Ŀ¼�¾Ϳ����ˡ�
* ʾ����Module\SysSPath
* Delphi ����������
```
 type
 { ֧�ֵ��ļ����� }
 TSPFileType = (ftDelphiDll, ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE);  
 procedure db_ShowDllForm_Plugins(
 var frm: TFormClass;
 var ft: TSPFileType;
 var strParentModuleName,
 strSubModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; 
 const bShow: Boolean = True); stdcall;
```

### VC��
* VC ԭ EXE ���̱��ֲ��䣬����õ� EXE�����Ʊ�������� OBJ��RES �ļ�����������Ŀ¼����
* �½�һ�� xxx.cpp �ļ������� db_ShowDllForm_Plugins �������������� xxx.obj��
* ���� xxx.obj ��ԭ���ı��� EXE ������ OBJ��RES �ļ����õ� Dll �ļ������õ� plugins Ŀ¼�¾Ϳ����ˡ�
* ʾ��1��DOC\VC\Dialog\7zip
* ʾ��2��DOC\VC\Dialog\Notepad2
* VC ����������
```
 enum TSPFileType {ftDelphiDll, ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE};
 extern "C" __declspec(dllexport) void db_ShowDllForm_Plugins(
 void** frm, 
 TSPFileType* spFileType,
 char** * strParentModuleName,
 char** strSubModuleName,
 char** strClassName,
 char** strWindowName,
 char** strIconFileName,
 const bool show = true)
```

## �ģ�Dll �����������˵��
```
 procedure db_ShowDllForm_Plugins(
 var frm: TFormClass;
 var ft: TSPFileType;
 var strParentModuleName,strSubModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar;
 const bShow: Boolean = True); stdcall;

 frm                 ��Delphi ר�á� Delphi �� Dll ������������VC �ÿգ�
 ft                  ���� Dll �����ͣ�֧�֣�
        ftDelphiDll   ��Delphi ���Ա�д�Ĵ��� Dll��
        ftVCDialogDll ��VC ����ʹ�� Dialog ��ʽ��д�Ĵ��� Dll��
        ftVCMFCDll    ��VC ����ʹ�� MFC    ��ʽ��д�Ĵ��� Dll��
        ftQTDll       ��QT ���Ա�д�Ĵ��� Dll��
        ftEXE         ��������� EXE �ļ���
 strParentModuleName ����ģ�����ƣ�
 strSubModuleName    ����ģ�����ƣ�
 strClassName        ��VC ר�ã�VC Dll ������������  Delphi �ÿգ�
 strWindowName       ��VC ר�ã�VC Dll �������������Delphi �ÿգ�
 strIconFileName     ��ͼ���ļ�����Ϊ�գ��� PBox �����У�ѡ��ͼ�ꣻ
 bShow               ��VC ר�ã��Ƿ���ʾ Dll ���壻��һ�ε��� VC Dll ʱ���ǲ��ô�����ʾ Dll ����ģ�ֻ��Ϊ�˻�ȡ������
```

## �壺��ɫ����
    ����֧�֣��˵���ʽ��ʾ����ť���Ի��򣩷�ʽ��ʾ���б��ӷ�ʽ��ʾ;
    PBox ��֧�ֽ�һ�� EXE ���������ʾ�����ǵĴ�����;
    ֧�ִ���������̬�仯�� EXE ����;
    ֧�� X86 ������� X64 ����, X64 ������� X86 ����;

## ����������������
    ������ݿ�֧�֣����ڱ��˶����ݿⲻ��Ϥ�����Կ�������������ҵ��ʱ�俪����;
    ��� VC MFC Dll / QT Dll �����֧��;

