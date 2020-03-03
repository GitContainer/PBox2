# PBox 是一个基于 Dll 动态库窗体的模块化开发平台

- [English](readme.md)

## 一：开发宗旨
    本着尽量少修改或不修改原有工程源代码的原则;
    支持 Delphi、VC、QT Dll窗体;

## 二：开发平台
    Delphi10.3.3、WIN10X64 下开发；
    代码没有使用任何第三方控件；
    引用了一些开源库，统一放置在 3rdparty 目录下面；直接引用，无需安装;
    WIN7X64、WIN10X64下测试通过；支持X86、X64;
    邮箱：dbyoung@sina.com
    QQ群：101611228

## 三：使用方法
### Delphi：
* Delphi 原 EXE 工程文件，修改为 Dll 工程。输出特定函数就可以了，原有代码不用作任何修改。
* 把编译后的 Dll 文件放置到 plugins 目录下就可以了。
* 示例：Module\SysSPath
* Delphi 函数声明：
```
 type
 { 支持的文件类型 }
 TSPFileType = (ftDelphiDll, ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE);  
 procedure db_ShowDllForm_Plugins(
 var frm: TFormClass;
 var ft: TSPFileType;
 var strParentModuleName,
 strSubModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; 
 const bShow: Boolean = True); stdcall;
```

### VC：
* VC 原 EXE 工程保持不变，编译得到 EXE。复制编译产生的 OBJ，RES 文件（复制整个目录）；
* 新建一个 xxx.cpp 文件，导出 db_ShowDllForm_Plugins 函数。编译生成 xxx.obj；
* 连接 xxx.obj 和原来的编译 EXE 产生的 OBJ，RES 文件，得到 Dll 文件，放置到 plugins 目录下就可以了。
* 示例1：DOC\VC\Dialog\7zip
* 示例2：DOC\VC\Dialog\Notepad2
* VC 函数声明：
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

## 四：Dll 输出函数参数说明
```
 procedure db_ShowDllForm_Plugins(
 var frm: TFormClass;
 var ft: TSPFileType;
 var strParentModuleName,strSubModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar;
 const bShow: Boolean = True); stdcall;

 frm                 ：Delphi 专用。 Delphi 中 Dll 主窗体类名；VC 置空；
 ft                  ：本 Dll 的类型；支持：
        ftDelphiDll   ：Delphi 语言编写的窗体 Dll；
        ftVCDialogDll ：VC 语言使用 Dialog 方式编写的窗体 Dll；
        ftVCMFCDll    ：VC 语言使用 MFC    方式编写的窗体 Dll；
        ftQTDll       ：QT 语言编写的窗体 Dll；
        ftEXE         ：带窗体的 EXE 文件；
 strParentModuleName ：父模块名称；
 strSubModuleName    ：子模块名称；
 strClassName        ：VC 专用；VC Dll 主窗体类名；  Delphi 置空；
 strWindowName       ：VC 专用；VC Dll 主窗体标题名；Delphi 置空；
 strIconFileName     ：图标文件；可为空，在 PBox 配置中，选择图标；
 bShow               ：VC 专用；是否显示 Dll 窗体；第一次调用 VC Dll 时，是不用创建显示 Dll 窗体的，只是为了获取参数。
```

## 五：特色功能
    界面支持，菜单方式显示、按钮（对话框）方式显示、列表视方式显示;
    PBox 还支持将一个 EXE 窗体程序显示在我们的窗体中;
    支持窗体类名动态变化的 EXE 程序;
    支持 X86 程序调用 X64 程序, X64 程序调用 X86 程序;

## 六：接下来工作：
    添加数据库支持（由于本人对数据库不熟悉，所以开发较慢，又是业余时间开发）;
    添加 VC MFC Dll / QT Dll 窗体的支持;

