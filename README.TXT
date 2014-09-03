基于as3xls及OpenOffice.org's Documentation of the Microsoft Excel File Format Excel Versions 2, 3, 4, 5, 95, 97, 2000, XP, 2003
FIXED:
1. 读取Excel文件-RECORD:
  1) +CODEPAGE，支持Sheet名称及Cell内容包含中文。
  2) +CONTINUE-SST。
  3) +DEFINEDNAME，支持公式中引用通过名称定义的单元格区域。
  4) +EXTERNALBOOK。
  5) +EXTERNALSHEET，支持诸如=Sheet1!A6+1.1公式中引用其他Sheet的单元格。
2. 读取Excel文件-FORMULA'S TOKEN:
  1) +tRef3dR/tRef3dV/tRef3dA，支持诸如=Sheet1!A6+1.1公式中引用其他Sheet的单元格。
  2) +tNameR/tNameV/tNameA。
  3) +tArea3dR/tArea3dV/tArea3dA。
  4) +tRefErrR/tRefErrV/tRefErrA。
  5) +tMissArg。
  6) +tAreaNR/tAreaNV/tAreaNA。
3. 读取Excel文件-FORMULA'S BUILTIN FUNCTION:
  1) +TRIM: 去除字符串首尾空格。
  2) +SUBSTITUTE: 替换字符串中的子串。
  3) +VLOOKUP: 查找值。
4. 目前在读取Excel文件时，对有公式的单元格，使用其缓存的结果值，仅当此值为NaN/undefined/null时才使用公式进行
        计算，从而可以避开公式引用依赖次序及循环引用的问题!!!

TODO:
1. 日期单元格的处理。
2. RECORD/TOKEN/BUILTIN FUNCTION进一步完善。
3. 多种BIFF格式的进一步支持。
4. Excel文件的修改及保存。

