package fxl.xls
{
  /**
   * Used to represent BIFF Records
   */
  public class Type
  {
    public static const DIMENSION:uint            = 0x00;
    public static const BLANK:uint                = 0x01;
    public static const INTEGER:uint              = 0x02;
    public static const NUMBER:uint               = 0x03;
    public static const LABEL:uint                = 0x04;
    public static const BOOLERR:uint              = 0x05;
    public static const FORMULA:uint              = 0x06;
    public static const STRING:uint               = 0x07;
    public static const ROW:uint                  = 0x08;
    public static const BOF:uint                  = 0x09;
    public static const EOF:uint                  = 0x0A;
    public static const INDEX:uint                = 0x0B;
    public static const CALCCOUNT:uint            = 0x0C;
    public static const CALCMODE:uint             = 0x0D;
    public static const PRECISION:uint            = 0x0E;
    public static const REFMODE:uint              = 0x0F;
    public static const DELTA:uint                = 0x10;
    public static const ITERATION:uint            = 0x11;
    public static const PROTECT:uint              = 0x12;
    public static const PASSWORD:uint             = 0x13;
    public static const HEADER:uint               = 0x14;
    public static const FOOTER:uint               = 0x15;
    public static const EXTERNALCOUNT:uint        = 0x16;
    public static const EXTERNALSHEET:uint        = 0x17;
    public static const DEFINEDNAME:uint          = 0x18;
    public static const WINDOWPROTECT:uint        = 0x19;
    public static const VERTICALPAGEBREAKS:uint   = 0x1A;
    public static const HORIZONTALPAGEBREAKS:uint = 0x1B;
    public static const NOTE:uint                 = 0x1C;
    public static const SELECTION:uint            = 0x1D;
    public static const FORMAT:uint               = 0x1E;
    public static const BUILTINFMTCOUNT:uint      = 0x1F;
    public static const COLUMNDEFAULT:uint        = 0x20;
    public static const ARRAY:uint                = 0x21;
    public static const DATEMODE:uint             = 0x22;
    public static const EXTERNALNAME:uint         = 0x23;
    public static const COLWIDTH:uint             = 0x24;
    public static const DEFAULTROWHEIGHT:uint     = 0x25;
    public static const LEFTMARGIN:uint           = 0x26;
    public static const RIGHTMARGIN:uint          = 0x27;
    public static const TOPMARGIN:uint            = 0x28;
    public static const BOTTOMMARGIN:uint         = 0x29;
    public static const PRINTHEADERS:uint         = 0x2A;
    public static const PRINTGRIDLINES:uint       = 0x2B;
    public static const FILEPASS:uint             = 0x2F;
    public static const FONT:uint                 = 0x31;
    public static const FONT2:uint                = 0x32;
    public static const DATATABLE:uint            = 0x36;
    public static const DATATABLE2:uint           = 0x37;
    public static const CONTINUE:uint             = 0x3C;
    public static const WINDOW1:uint              = 0x3D;
    public static const WINDOW2:uint              = 0x3E;
    public static const BACKUP:uint               = 0x40;
    public static const PANE:uint                 = 0x41;
    public static const CODEPAGE:uint             = 0x42;
    public static const XF:uint                   = 0x43;
    public static const IXFE:uint                 = 0x44;
    public static const FONTCOLOR:uint            = 0x45;
    public static const PLS:uint                  = 0x4D;
    public static const DCONREF:uint              = 0x51;
    public static const DEFAULTCOLWIDTH:uint      = 0x55;
    // public static const BUILTINFMTCOUNT:uint   = 0x56;
    public static const XCT:uint                  = 0x59;
    public static const CRN:uint                  = 0x5A;
    public static const FILESHARING:uint          = 0x5B;
    public static const WRITEACCESS:uint          = 0x5C;
    public static const UNCALCED:uint             = 0x5E;
    public static const SAVERECALC:uint           = 0x5F;
    public static const OBJECTPROTECT:uint        = 0x63;
    public static const COLINFO:uint              = 0x7D;
    public static const GUTS:uint                 = 0x80;
    public static const SHEETPR:uint              = 0x81;
    public static const GRIDSET:uint              = 0x82;
    public static const HCENTER:uint              = 0x83;
    public static const VCENTER:uint              = 0x84;
    public static const SHEET:uint                = 0x85;
    public static const WRITEPROT:uint            = 0x86;
    public static const COUNTRY:uint              = 0x8C;
    public static const HIDEOBJ:uint              = 0x8D;
    public static const SORT:uint                 = 0x90;
    public static const PALETTE:uint              = 0x92;
    public static const STANDARDWIDTH:uint        = 0x99;
    public static const SCL:uint                  = 0xA0;
    public static const PAGESETUP:uint            = 0xA1;
    public static const GCW:uint                  = 0xAB;
    public static const MULRK:uint                = 0xBD;
    public static const MULBLANK:uint             = 0xBE;
    public static const RSTRING:uint              = 0xD6;
    public static const DBCELL:uint               = 0xD7;
    public static const BOOKBOOL:uint             = 0xDA;
    public static const SCENPROTECT:uint          = 0xDD;
    // public static const XF:uint                = 0xE0;
    public static const MERGEDCELLS:uint          = 0xE5;
    public static const BITMAP:uint               = 0xE9;
    public static const PHONETICPR:uint           = 0xEF;
    public static const SST:uint                  = 0xFC;
    public static const LABELSST:uint             = 0xFD;
    public static const EXTSST:uint               = 0xFF;
    public static const LABELRANGES:uint          = 0x15F;
    public static const USESELFS:uint             = 0x160;
    public static const DSF:uint                  = 0x161;
    public static const EXTERNALBOOK:uint         = 0x1AE;
    public static const CFHEADER:uint             = 0x1B0;
    public static const CFRULE:uint               = 0x1B1;
    public static const DATAVALIDATIONS:uint      = 0x1B2;
    public static const HYPERLINK:uint            = 0x1B8;
    public static const DATAVALIDATION:uint       = 0x1BE;
    // public static const DIMENSION:uint         = 0x200;
    // public static const BLANK:uint             = 0x201;
    // public static const NUMBER:uint            = 0x203;
    // public static const LABEL:uint             = 0x204;
    // public static const BOOLERR:uint           = 0x205;
    // public static const FORMULA:uint           = 0x206;
    // public static const STRING:uint            = 0x207;
    // public static const ROW:uint               = 0x208;
    // public static const BOF:uint               = 0x209;
    // public static const INDEX:uint             = 0x20B;
    // public static const DEFINEDNAME:uint       = 0x218;
    // public static const ARRAY:uint             = 0x221;
    // public static const EXTERNALNAME:uint      = 0x223;
    // public static const DEFAULTROWHEIGHT:uint  = 0x225;
    // public static const FONT:uint              = 0x231;
    // public static const DATATABLE:uint         = 0x236;
    // public static const WINDOW2:uint           = 0x23E;
    // public static const XF:uint                = 0x243;
    public static const RK:uint                   = 0x27E;
    public static const STYLE:uint                = 0x293;
    // public static const FORMULA:uint           = 0x406;
    // public static const BOF:uint               = 0x409;
    // public static const FORMAT:uint            = 0x41E;
    // public static const XF:uint                = 0x443;
    public static const SHAREDFMLA:uint           = 0x4BC;
    public static const QUICKTIP:uint             = 0x800;
    // public static const BOF:uint                  = 0x809;
    public static const SHEETLAYOUT:uint          = 0x862;
    public static const SHEETPROTECTION:uint      = 0x867;
    public static const RANGEPROTECTION:uint      = 0x868;

  }
}