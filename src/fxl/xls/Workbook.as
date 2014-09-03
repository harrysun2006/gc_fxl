package fxl.xls
{
  import flash.utils.ByteArray;
  import flash.utils.Endian;

  // import fxl.Util;
  import fxl.biff.BIFFReader;
  import fxl.biff.BIFFVersion;
  import fxl.biff.BIFFWriter;
  import fxl.biff.Record;
  import fxl.cdf.CDFReader;
  import fxl.xls.formula.Formula;
  import fxl.xls.style.XFormat;

  import mx.collections.ArrayCollection;

  public class Workbook
  {
    public static const BASE1899:uint = 0;
    public static const BASE1904:uint = 1;

    private var _name:String;
    private var br:BIFFReader;
    private var _version:uint;
    private var dateMode:uint;
    private var _codepage:uint;
    private var _charset:String;

    private var globalFormats:Array = new Array();
    private var globalXFormats:Array = new Array();
    private var notes:Array;

    private var currentSheet:Worksheet;
    private var currentSheetIdx:uint = 0;
    private var _sheets:ArrayCollection = new ArrayCollection();
    private var _isheets:Array = new Array();
    private var _books:Array = new Array();
    private var _names:Array = new Array();

    private var lastRecordType:uint;

    private var _sst:Array;
    private var _sstCount:uint;
    private var numWorkbookStrings:uint;
    private var rrest:Object;

    private const handlers:Array = initHandlers();
    private static const ignore:Array = [
      0x4D, 0xE1, 0xC0, 0xC1, 0xE1, 
      0xE2, 0x5D, 0x9C, 0xBF, 0xEB, 
      0xEE, 0xF1, 0x13D, 0x1AF, 0x1B6, 
      0x1B7, 0x1BC, 0x1C0, 0x1C1,	0x1C2, 
      0x863, 0x8C8, 0x105C];
    private static const CODEPAGES:Object = {
        367:"ASCII",
        932:"SHIFT_JIS",
        936:"GBK",
        950:"BIG5",
        1200:"UTF-16",
        1201:"UNICODEFFFE",
        1251:"ISO-8859-1",
        65000:"UTF-7",
        65001:"UTF-8"};
    private static const RE1:RegExp = new RegExp(String.fromCharCode(0x03), "g");

    public function Workbook(n:String = "")
    {
      _name = n;
    }

    protected function initHandlers():Array
    {
      var handlers:Array = [
        DIMENSION, 
        BLANK,              INTEGER,              NUMBER,         LABEL,           BOOLERR, 
        FORMULA,            STRING,               ROW,            BOF,             EOF,             // 10
        INDEX,              CALCCOUNT,            CALCMODE,       PRECISION,       REFMODE,         
        DELTA,              ITERATION,            PROTECT,        PASSWORD,        HEADER,          // 20
        FOOTER,             EXTERNALCOUNT,        EXTERNALSHEET,  DEFINEDNAME,     WINDOWPROTECT,   
        VERTICALPAGEBREAKS, HORIZONTALPAGEBREAKS, NOTE,           SELECTION,       FORMAT,          // 30
        BUILTINFMTCOUNT,    COLUMNDEFAULT,        ARRAY,          DATEMODE,        EXTERNALNAME,    
        COLWIDTH,           DEFAULTROWHEIGHT,     LEFTMARGIN,     RIGHTMARGIN,     TOPMARGIN,       // 40
        BOTTOMMARGIN,       PRINTHEADERS,         PRINTGRIDLINES, null,            null,            
        null,               FILEPASS,             null,           FONT,            FONT2,           // 50
        null,               null,                 null,           DATATABLE,       DATATABLE2,        
        null,               null,                 null,           null,            CONTINUE,        // 60 
        WINDOW1,            WINDOW2,              null,           BACKUP,          PANE,            
        CODEPAGE,           XF,                   IXFE,           FONTCOLOR,       null,            // 70
        null,               null,                 null,           null,            null,            
        null,               PLS,                  null,           null,            null,            // 80
        DCONREF,            null,                 null,           null,            DEFAULTCOLWIDTH,    
        BUILTINFMTCOUNT,    null,                 null,           XCT,             CRN,             // 90
        FILESHARING,        WRITEACCESS,          null,           UNCALCED,        SAVERECALC,      
        null,               null,                 null,           OBJECTPROTECT,   null,            // 100
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 110 
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 120
        null,               null,                 null,           null,            COLINFO,         
        null,               null,                 GUTS,           SHEETPR,         GRIDSET,         // 130
        HCENTER,            VCENTER,              SHEET,          WRITEPROT,       null,            
        null,               null,                 null,           null,            COUNTRY,         // 140
        HIDEOBJ,            null,                 null,           SORT,            null,            
        PALETTE,            null,                 null,           null,            null,            // 150
        null,               null,                 STANDARDWIDTH,  null,            null,            
        null,               null,                 null,           null,            SCL,             // 160
        PAGESETUP,          null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 170
        GCW,                null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 180
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           MULRK,           MULBLANK,        // 190
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 200
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 210
        null,               null,                 null,           RSTRING,         DBCELL,          
        null,               null,                 BOOKBOOL,       null,            null,            // 220
        SCENPROTECT,        null,                 null,           XF,              null,            
        null,               null,                 null,           MERGEDCELLS,     null,            // 230
        null,               null,                 BITMAP,         null,            null,            
        null,               null,                 null,           PHONETICPR,      null,            // 240
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 250
        null,               SST,                  LABELSST,       null,            EXTSST,          
        null,               null,                 null,           null,            null,            // 260
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 270
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 280 
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 290
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 300
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 310
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 320
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 330
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 340
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 350
        LABELRANGES,        USESELFS,             DSF,            null,            null,            
        null,               null,                 null,           null,            null,            // 360
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 370
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 380
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 390
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 400
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 410
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 420
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            EXTERNALBOOK,    // 430
        null,               CFHEADER,             CFRULE,         DATAVALIDATIONS, null,            
        null,               null,                 null,           null,            HYPERLINK,       // 440
        null,               null,                 null,           null,            null,            
        DATAVALIDATION,     null,                 null,           null,            null,            // 450
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 460
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 470
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 480
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 490
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 500
        null,               null,                 null,           null,            null,            
        null,               null,                 null,           null,            null,            // 510
        null,               DIMENSION,            BLANK,          null,            NUMBER,          
        LABEL,              BOOLERR,              FORMULA,        STRING,          ROW,             // 520
        BOF,                null,                 INDEX,          null,            null
        ];
      handlers[0x221] = ARRAY;
      handlers[0x225] = DEFAULTROWHEIGHT;
      handlers[0x231] = FONT;
      handlers[0x23E] = WINDOW2;
      handlers[0x243] = XF;
      handlers[Type.RK] = RK;
      handlers[Type.STYLE] = STYLE;
      handlers[Type.SHAREDFMLA] = SHAREDFMLA;
      handlers[Type.QUICKTIP] = QUICKTIP;
      handlers[Type.SHEETLAYOUT] = SHEETLAYOUT;
      handlers[Type.SHEETPROTECTION] = SHEETPROTECTION;
      handlers[Type.RANGEPROTECTION] = RANGEPROTECTION;
      handlers[0x406] = FORMULA;
      handlers[0x409] = BOF;
      handlers[0x41E] = FORMAT;
      handlers[0x443] = XF;
      handlers[0x809] = BOF;
      return handlers;
    }

    public function get version():uint
    {
      return _version;
    }

    public function get codepage():uint
    {
      return _codepage;
    }

    public function get charset():String
    {
      return _charset;
    }

    public function get sheets():ArrayCollection
    {
      return _sheets;
    }

    public function sheet(index:uint):Worksheet
    {
      var s:Worksheet = null, s1:uint;
      if (_version == BIFFVersion.BIFF8 && index < _isheets.length)
      {
        var o:Object = _isheets[index];
        s = (o && o.b == 0 && o.s1 < _sheets.length) ? _sheets.getItemAt(o.s1) as Worksheet : null;
      }
      return s;
    }

    public function name(index:uint):Object
    {
      return _names[index-1];
    }

    /**
     * Saves the first sheet in the sheets array as a BIFF2 document. Saving formulas
     * is not currently supported.
     * @return A ByteArray containing the saved sheet in BIFF2 form
     *
     */
    public function save():ByteArray
    {
      var s:Worksheet = _sheets[0] as Worksheet;
      var br:BIFFWriter = new BIFFWriter();

      // Write the BOF and header records
      var bof:Record = new Record(Type.BOF);
      bof.data.writeShort(BIFFVersion.BIFF2);
      bof.data.writeByte(0);
      bof.data.writeByte(0x10);
      br.writeTag(bof);

      // Date mode
      var dateMode:Record = new Record(Type.DATEMODE);
      dateMode.data.writeShort(1);
      br.writeTag(dateMode);

      // Store built in formats
      var formats:Array = ["General", 
        "0", "0.00", "#,##0", "#,##0.00", 
        "", "", "", "",
        "0%", "0.00%", "0.00E+00",
        "#?/?", "#??/??",
        "M/D/YY", "D-MMM-YY", "D-MMM", "MMM-YY"];

      var numfmt:Record = new Record(Type.BUILTINFMTCOUNT);
      numfmt.data.writeShort(formats.length);
      br.writeTag(numfmt);

      for (var n:uint = 0; n < formats.length; n++)
      {
        var fmt:Record = new Record(Type.FORMAT);
        fmt.data.writeByte(formats[n].length);
        fmt.data.writeUTFBytes(formats[n]);
        br.writeTag(fmt);
      }

      var dimensions:Record = new Record(Type.DIMENSION);
      dimensions.data.writeShort(0);
      dimensions.data.writeShort(s.rows+1);
      dimensions.data.writeShort(0);
      dimensions.data.writeShort(s.cols+1);
      br.writeTag(dimensions);

      var b:ByteArray;
      for (var r:uint = 0; r < s.rows; r++)
      {
        for (var c:uint = 0; c < s.cols; c++)
        {
          var value:* = s.cell(r, c).value;
          var cell:Record = new Record(1);
          cell.data.writeShort(r);
          cell.data.writeShort(c);

          if (value is Date)
          {
            var dateNum:Number = (value.time / 86400000) + 24106.667;
            cell.type = Type.NUMBER;
            cell.data.writeByte(0);
            cell.data.writeByte(15);
            cell.data.writeByte(0);
            cell.data.writeDouble(dateNum);
          }
          else if (isNaN(Number(value)) == false && String(value) != "")
          {
            cell.type = Type.NUMBER;
            cell.data.writeByte(0);
            cell.data.writeByte(0);
            cell.data.writeByte(0);
            cell.data.writeDouble(value);
          }
          else if (String(value).length > 0)
          {
            cell.type = Type.LABEL;
            cell.data.writeByte(0);
            cell.data.writeByte(0);
            cell.data.writeByte(0);
            b = new ByteArray();
            b.writeMultiByte(value, "GBK");
            cell.data.writeByte(b.length);
            // cell.data.writeUTFBytes(value);
            cell.data.writeBytes(b);
          }
          else
          {
            cell.type = Type.BLANK;
            cell.data.writeByte(0);
            cell.data.writeByte(0);
            cell.data.writeByte(0);
          }

          br.writeTag(cell);
        }
      }

      // Finally, the closing EOF record
      var eof:Record = new Record(Type.EOF);
      br.writeTag(eof);

      br.stream.position = 0;
      return br.stream;
    }

    /**
     * Loads the sheets from a ByteArray containing an Excel file. If the ByteArray contains a CDF file the Workbook stream
     * will be extracted and loaded.
     *
     * @see fxl.cdf.CDFReader
     */
    public function load(xls:ByteArray):void
    {
      // Newer workbooks are actually cdf files which must be extracted
      if (CDFReader.isCDFFile(xls))
      {
        var cdf:CDFReader = new CDFReader(xls);
        xls = cdf.loadDirectory("Workbook");
        if (xls == null)
          xls = cdf.loadDirectoryEntry(0);
      }
      _sst = new Array();
      br = new BIFFReader(xls);
      // data:之前剩余字节数组, n1:字符串v长度, f1:单/双字节, e1:v被截断, e3:f3+n3被截断
      rrest = {data:new ByteArray(), n1:0, f1:0, e1:false, e3:false, e4:false};

      var unknown:Array = [];
      var r:Record, n:String;
      while ((r = br.readTag()) != null)
      {
        if (ignore.indexOf(r.type) != -1)
        {
          continue;
        }
        if (r.type != Type.CONTINUE)
        {
          lastRecordType = r.type;
        }
        if (handlers[r.type] is Function)
        {
          (handlers[r.type] as Function).call(this, r, currentSheet);
        }
        else
        {
          /*
             n = currentSheet ? currentSheet.name : "";
             Util.debugByteArray(r.data, br.position+", sheet: "+n+", r: "+r.type+", "+r.length);
           */
          unknown.push(r.type);
        }
      }
      trace(_name + "[SST count: " + _sstCount + ", _sst.length: " + _sst.length + "]");
      if (unknown.length > 0)
      {
        //throw new Error("Unsupported BIFF records: " + unknown.join(", "));
        trace("Unsupported BIFF records: " + unknown.join(", "));
      }
    }

    private function readString(b:ByteArray):String
    {
      var n1:uint = b.readUnsignedShort(); // Length of the string (character count, ln)
      var opts:uint = b.readByte();
      var f1:Boolean = (opts & 0x01) == 0x01; // Character compression (ccompr): 0=Compressed(8-bit); 1=Uncompressed(16-bit)
      var f3:Boolean = (opts & 0x04) == 0x04; // Asian phonetic settings (phonetic): 0=Does not contain Asian phonetic settings; 1=Contains Asian phonetic settings
      var f4:Boolean = (opts & 0x08) == 0x08; // Rich-Text settings (richtext): 0=Does not contain Rich-Text settings; 1=Contains Rich-Text settings
      var nn1:uint, n4:uint, n3:uint, v:String;
      n4 = f4 ? b.readUnsignedShort() : 0; // (optional, only if richtext=1) Number of Rich-Text formatting runs (rt)
      n3 = f3 ? b.readUnsignedInt() : 0; // (optional, only if phonetic=1) Size of Asian phonetic settings block (in bytes, sz)
      // Character array (8-bit characters or 16-bit characters, dependent on ccompr)
      nn1 = f1 ? 2*n1 : n1;
      rrest.e1 = rrest.e3 = rrest.e4 = false;
      if (nn1 > b.bytesAvailable)
      {
        rrest.e1 = true;
        return null;
      }
      v = f1 ? b.readMultiByte(nn1, _charset) : b.readUTFBytes(nn1);
      if (4 * n4 > b.bytesAvailable)
      {
        rrest.e4 = true;
        return null;
      }
      b.position += 4*n4; // (optional, only if richtext=1) List of rt formatting runs (->2.5.1)
      if (f3) // (optional, only if phonetic=1) Asian Phonetic Settings Block (see below)
      {
        if (n3 > b.bytesAvailable)
        {
          rrest.e3 = true;
          return null;
        }
        b.position += n3;
      }
      return v;
    }

    private function readRK(r:Record, s:Worksheet):Number
    {
      var raw:uint = r.data.readUnsignedInt();
      var div100:Boolean = (raw & 0x00000001) == 1
      var intVal:Boolean = (raw & 0x00000002) == 2;

      r.data.position -= 4;

      var value:Number;
      if (intVal)
      {
        value = r.data.readInt() >> 2;
      }
      else
      {
        var b:ByteArray = new ByteArray();
        b[7] = 0;
        b[6] = 0;
        b[5] = 0;
        b[4] = 0;
        b[3] = r.data.readUnsignedByte();
        b[2] = r.data.readUnsignedByte();
        b[1] = r.data.readUnsignedByte();
        b[0] = r.data.readUnsignedByte();
        value = b.readDouble();
      }

      if (div100)
      {
        value = Math.round(value) / 100;
      }

      return value;
    }

    private function ARRAY(r:Record, s:Worksheet):void
    {
      var firstRow:uint = r.data.readUnsignedShort();
      var lastRow:uint = r.data.readUnsignedShort();
      var firstCol:uint = r.data.readUnsignedByte();
      var lastCol:uint = r.data.readUnsignedByte();

      var alwaysRecalculate:Boolean;
      if (_version == BIFFVersion.BIFF2)
      {
        alwaysRecalculate = r.data.readUnsignedByte() == 0x01;
      }
      else
      {
        alwaysRecalculate = (r.data.readUnsignedShort() & 0x0001) == 0x0001;
      }

      if (_version == BIFFVersion.BIFF8)
      {
        r.data.position += 4;
      }

      var tokens:ByteArray = new ByteArray();
      r.data.readBytes(tokens, 0, r.data.bytesAvailable);
    }

    private function BACKUP(r:Record, s:Worksheet):void
    {
      var saveBackup:Boolean = r.data.readUnsignedShort() == 1;
    }

    private function BITMAP(r:Record, s:Worksheet):void
    {
    /*
       var n:String = currentSheet ? currentSheet.name : "";
       Util.debugByteArray(r.data, "BITMAP@"+n);
     */
    }

    private function BLANK(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();
      s.setCell(row, col, "");
    }

    private function BOF(r:Record, s:Worksheet):void
    {
      var v:uint = r.data.readUnsignedShort();
      var type:uint = r.data.readUnsignedShort();
      if (type == Worksheet.T_SHEET)
      {
        if (_sheets.length == 0)
        {
          var newSheet:Worksheet = new Worksheet(this);
          newSheet.dateMode = dateMode;
          _sheets.addItem(newSheet);
          newSheet.formats = newSheet.formats.concat(globalFormats);
          newSheet.xformats = newSheet.xformats.concat(globalXFormats);
        }
        notes = new Array();
        currentSheet = _sheets.getItemAt(currentSheetIdx) as Worksheet;
        currentSheet.type = type;
        currentSheetIdx++;
      }
      if (r.type == 0x9)
      {
        _version = BIFFVersion.BIFF2
      }
      else if (r.type == 0x209)
      {
        _version = BIFFVersion.BIFF3;
      }
      else if (r.type == 0x409)
      {
        _version = BIFFVersion.BIFF4;
      }
      else if (r.type == 0x809 && v == 0x0500)
      {
        _version = BIFFVersion.BIFF5;
      }
      else if (r.type == 0x809 && v == 0x0600)
      {
        _version = BIFFVersion.BIFF8;
      }
      // trace("_version="+this._version);
    }

    private function BOOKBOOL(r:Record, s:Worksheet):void
    {
      var options:uint = r.data.readUnsignedShort();
    }

    private function BOOLERR(r:Record, s:Worksheet):void
    {
    }

    private function BOTTOMMARGIN(r:Record, s:Worksheet):void
    {
      var size:Number = r.data.readDouble();
    }

    private function BUILTINFMTCOUNT(r:Record, s:Worksheet):void
    {
      var numBuildInFormats:uint = r.data.readUnsignedShort();
    }

    private function CALCCOUNT(r:Record, s:Worksheet):void
    {
      var iterations:uint = r.data.readUnsignedShort();
    }

    private function CALCMODE(r:Record, s:Worksheet):void
    {
      var mode:uint = r.data.readUnsignedShort();
    }

    private function CFHEADER(r:Record, s:Worksheet):void
    {
    }

    private function CFRULE(r:Record, s:Worksheet):void
    {
    }

    private function CODEPAGE(r:Record, s:Worksheet):void
    {
      _codepage = r.data.readUnsignedShort();
      _charset = CODEPAGES.hasOwnProperty(_codepage) ? CODEPAGES[_codepage] : "UTF-8";
      // trace("codepage="+_codepage);
    }

    private function COLINFO(r:Record, s:Worksheet):void
    {
    }

    private function COLUMNDEFAULT(r:Record, s:Worksheet):void
    {
    }

    private function COLWIDTH(r:Record, s:Worksheet):void
    {
    }

    /**
     * SST及CONTINUE-SST记录中在末尾的v或f3+n3可能会被截断
     * 如果v被截断, 则下一个CONTINUE-SST记录的长度后第一个字节为压缩码(f1)
     * 如3C 00 47 06 01(3C 00是CONTINUE记录类型码)中长度(47 06)之后的00/01表示字符串是单/双字节
     * 如果v未被截断, f3+n3被截断, 则下一个CONTINUE-SST记录长度之后没有压缩码(f1)
     **/
    private function CONTINUE(r:Record, s:Worksheet):void
    {
      switch (lastRecordType)
      {
        case 0xEC:
          var flags:uint = r.data.readUnsignedByte();
          var note:String = r.data.readUTFBytes(r.data.bytesAvailable);
          notes.push(note);
          break;
        case 0xFC:
          // trace("pos@stream: " + br.position);
          if (rrest.data.length > 0)
          {
            // Util.debugByteArray(rrest.data, "rrest");
            // Util.debugByteArray(r.data, "1@trim");
            if (rrest.e1)
              r.ltrim();
            // Util.debugByteArray(r.data, "2@trim");
            r.insert(rrest.data);
              // Util.debugByteArray(r.data, "1@insert");
          }
          SST2(r, s);
          // trace(_sst.length > 0 ? _sst[_sst.length - 1] : "");
          break;
        default:
          /*
             var n:String = currentSheet ? currentSheet.name : "";
             Util.debugByteArray(r.data, "CONTINUE("+lastRecordType+")@"+n);
           */
          break;
      }
    }

    private function COUNTRY(r:Record, s:Worksheet):void
    {
      var excelCountry:uint = r.data.readUnsignedShort();
      var systemCountry:uint = r.data.readUnsignedShort();
    }

    private function CRN(r:Record, s:Worksheet):void
    {
    }

    private function DATATABLE(r:Record, s:Worksheet):void
    {
    }

    private function DATATABLE2(r:Record, s:Worksheet):void
    {
    }

    private function DATAVALIDATION(r:Record, s:Worksheet):void
    {
    }

    private function DATAVALIDATIONS(r:Record, s:Worksheet):void
    {
    }

    private function DATEMODE(r:Record, s:Worksheet):void
    {
      var baseDate:uint = r.data.readUnsignedShort();
      dateMode = baseDate;
      for (var n:uint = 0; n < _sheets.length; n++)
      {
        _sheets.getItemAt(n).dateMode = dateMode;
      }
    }

    private function DBCELL(r:Record, s:Worksheet):void
    {
    }

    private function DCONREF(r:Record, s:Worksheet):void
    {
    }

    private function DEFAULTROWHEIGHT(r:Record, s:Worksheet):void
    {
      var defRowHeight:Number = r.data.readUnsignedShort();
    }

    private function DEFAULTCOLWIDTH(r:Record, s:Worksheet):void
    {
      var defColWidth:uint = r.data.readUnsignedShort();
    }

    private function DEFINEDNAME(r:Record, s:Worksheet):void
    {
      var flag:uint, shortcut:uint;
      var ln:uint; // length of the name
      var sz:uint; // size of the formula data
      var indexToSheet:uint; // 0 = Global name, otherwise index to sheet (one-based)
      var lm:uint; // length of menu text
      var ld:uint; // length of description text
      var lh:uint; // length of help topic text
      var ls:uint; // length of statusbar text
      var name:String;
      var tokens:ByteArray = new ByteArray();
      var formula:Formula = null;
      if (_version == BIFFVersion.BIFF8)
      {
        flag = r.data.readUnsignedShort();
        shortcut = r.data.readUnsignedByte();
        ln = r.data.readUnsignedByte();
        sz = r.data.readUnsignedShort();
        r.data.position += 2;
        indexToSheet = r.data.readUnsignedShort();
        lm = r.data.readUnsignedByte();
        ld = r.data.readUnsignedByte();
        lh = r.data.readUnsignedByte();
        ls = r.data.readUnsignedByte();
        r.data.position += 1;
        name = r.data.readUTFBytes(ln);
        if (sz > 0)
        {
          // Util.debugByteArray(tokens, name+"@DEFINEDNAME");
          r.data.readBytes(tokens, 0, sz);
        }
        formula = new Formula(this, s, 0, 0, tokens);
        _names.push({name:name, formula:formula});
      }
    }

    private function DELTA(r:Record, s:Worksheet):void
    {
      var delta:Number = r.data.readDouble();
    }

    private function DIMENSION(r:Record, s:Worksheet):void
    {
      // For some reason sometimes the dimension record is blank. Reading it fails in this case
      if (r.data.length == 0)
      {
        return;
      }

      // Using the biff _version doesn't seem to work; instead use the record size to figure out
      // whether the row indeces are ints or shorts;
      var firstRow:uint = r.length == 14 ? r.data.readUnsignedInt() : r.data.readUnsignedShort();
      var lastRow:uint = r.length == 14 ? r.data.readUnsignedInt() : r.data.readUnsignedShort();
      var firstCol:uint = r.data.readUnsignedShort();
      var lastCol:uint = r.data.readUnsignedShort();
      s.resize(lastRow, lastCol);
    }

    private function DSF(r:Record, s:Worksheet):void
    {
    }

    private function EOF(r:Record, s:Worksheet):void
    {
    }

    private function EXTERNALBOOK(r:Record, s:Worksheet):void
    {
      var nm:uint = r.data.readUnsignedShort(), i:uint;
      var t:uint = r.data.readUnsignedShort();
      if (t == 0x0401)
        return; // own book
      else
        r.data.position -= 2;
      var nb:String = readString(r.data), ns:String, cc:Array=[];
      i = 0;
      while (i < 3)
      {
        cc[i] = nb.length > i ? nb.charCodeAt(i) : 0;
        i++;
      }
      if (cc[0] == 0x01 && cc[1] == 0x01 && cc[2] == 0x40)
        nb = "\\\\"+nb.substring(3);
      else if (cc[0] == 0x01 && cc[1] < 26)
        nb = String.fromCharCode(cc[1]+0x41)+":\\"+nb.substring(2);
      else if (cc[0] == 0x01)
        nb = nb.substring(1);
      nb = nb.replace(RE1, "\\");
      var b:Object = {nb:nb, ns:[]};
      for (i = 0; i < nm; i++)
      {
        ns = readString(r.data);
        b.ns.push(ns);
      }
      _books.push(b);
    }

    private function EXTERNALNAME(r:Record, s:Worksheet):void
    {
    }

    private function EXTERNALCOUNT(r:Record, s:Worksheet):void
    {
    }

    private function EXTERNALSHEET(r:Record, s:Worksheet):void
    {
      if (_version == BIFFVersion.BIFF8)
      {
        var c:uint = r.data.readUnsignedShort();
        var i:uint = 0;
        var b:uint, s1:uint, s2:uint;
        for (i = 0; i < c; i++)
        {
          b = r.data.readUnsignedShort();
          s1 = r.data.readUnsignedShort();
          s2 = r.data.readUnsignedShort();
          _isheets.push({b:b, s1:s1, s2:s2});
        }
      }
    }

    private function EXTSST(r:Record, s:Worksheet):void
    {
    }

    private function FILEPASS(r:Record, s:Worksheet):void
    {
    }

    private function FILESHARING(r:Record, s:Worksheet):void
    {
    }

    private function FONT(r:Record, s:Worksheet):void
    {
      var height:Number = r.data.readUnsignedShort();
      var attributes:uint = r.data.readUnsignedShort();

      if (r.type == 0x231 || _version >= BIFFVersion.BIFF5)
      {
        var colorIndex:uint = r.data.readUnsignedShort();
      }

      var len:uint;
      var name:String;
      if (r.type == 0x231 && _version <= BIFFVersion.BIFF4)
      {
        len = r.data.readUnsignedByte();
        name = r.data.readUTFBytes(len);
      }
    }

    private function FONT2(r:Record, s:Worksheet):void
    {
    }

    private function FONTCOLOR(r:Record, s:Worksheet):void
    {
      var color:uint = r.data.readUnsignedShort();
    }

    private function FOOTER(r:Record, s:Worksheet):void
    {
      if (r.data.bytesAvailable == 0)
      {
        return;
      }
      var len:uint = r.data.readUnsignedByte();
      var string:String = r.data.readUTFBytes(len).substr(2); // Skip two bytes b/c of commands (left, center, etc)
      s.footer = string;
    }

    /**
     * Sample:
     * 05 00 13 00 01 22 00 E5 FF 22 00 23 00 2C 00 23
     * 00 23 00 30 00 3B 00 22 00 E5 FF 22 00 5C 00 2D
     * 00 23 00 2C 00 23 00 23 00 30 00
     * 05 00 - flag/option ?
     * 13 00 - len
     * 01    - single/multi ?
     * ...   - format
     **/
    private function FORMAT(r:Record, s:Worksheet):void
    {
      var id:uint = r.data.readUnsignedShort();
      var len:uint = _version <= BIFFVersion.BIFF5 ? r.data.readUnsignedByte() : r.data.readUnsignedShort();
      var multi:uint = _version <= BIFFVersion.BIFF5 ? 0 : r.data.readByte();

      // var str:String = r.data.readUTFBytes(len);
      var str:String = multi ? r.data.readMultiByte(2*len, _charset) : r.data.readUTFBytes(len);
      // trace("format: len="+len+", f="+str);
      if (s is Worksheet)
      {
        s.formats.push(str);
      }
      else
      {
        globalFormats.push(str);
      }
    }

    private function FORMULA(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();
      var indexToXF:uint;

      // Cell attributes
      if (_version == BIFFVersion.BIFF2)
      {
        r.data.readUnsignedByte();
        r.data.readUnsignedByte();
        r.data.readUnsignedByte();
      }
      else
      {
        indexToXF = r.data.readUnsignedShort();
      }

      var result:Number = r.data.readDouble();
      var alwaysRecalculate:Boolean, calculateOnOpen:Boolean, partOfSharedFormula:Boolean;
      var t1:uint;
      if (_version == BIFFVersion.BIFF2)
      {
        alwaysRecalculate = r.data.readUnsignedByte() == 1;
      }
      else
      {
        t1 = r.data.readUnsignedShort();
        alwaysRecalculate = (t1 & 0x0001) == 0x01;
        calculateOnOpen = (t1 & 0x0002) == 0x02;
        partOfSharedFormula = (t1 & 0x0008) == 0x08;
      }

      if (_version >= BIFFVersion.BIFF5)
      {
        // For some reason in BIFF5-8 there are 4 unused bytes before the token array
        r.data.position += 4;
      }

      var tokenArrSize:uint = (_version == BIFFVersion.BIFF2) ? r.data.readUnsignedByte() : r.data.readUnsignedShort();
      var tokens:ByteArray = new ByteArray();
      r.data.readBytes(tokens, 0, tokenArrSize);

      var f:Formula = new Formula(this, s, row, col, tokens, result);
      s.setCell(row, col, f);
      // trace("formula("+row+","+col+")="+result);

      var fmt:String = s.formats[s.xformats[indexToXF].format];
      if (fmt == null || fmt.length == 0)
      {
        fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
      }

      s.cell(row, col).format = fmt;
    }

    // Global Column Width
    private function GCW(r:Record, s:Worksheet):void
    {
      var bitfieldSize:uint = r.data.readUnsignedShort();
    }

    private function GRIDSET(r:Record, s:Worksheet):void
    {
      var printGridLinesOptionEverChanged:Boolean = r.data.readUnsignedByte() == 1;
    }

    private function GUTS(r:Record, s:Worksheet):void
    {
      var rowOutlineWidth:uint = r.data.readUnsignedShort();
      var colOutlineHeight:uint = r.data.readUnsignedShort();
      var visibleRowLevels:uint = r.data.readUnsignedShort();
      var visibleColLevels:uint = r.data.readUnsignedShort();
    }

    private function HCENTER(r:Record, s:Worksheet):void
    {
      // 0 = left align, 1 = centered
      var center:uint = r.data.readUnsignedShort();
    }

    private function HEADER(r:Record, s:Worksheet):void
    {
      var name:String;
      if (r.data.length == 0)
      {
        name = "";
      }
      else if (_version == BIFFVersion.BIFF8)
      {
        // string = r.readUnicodeStr16();
        name = readString(r.data);
      }
      else
      {
        var len:uint = r.data.readUnsignedByte();
        name = r.data.readUTFBytes(len);
      }
      s.header = name;
    }

    private function HIDEOBJ(r:Record, s:Worksheet):void
    {
      /*
         0 = Show objects
         1 = Show placeholders
         2 = Hide objects
       */
      var viewingMode:uint = r.data.readUnsignedShort();
    }

    private function HORIZONTALPAGEBREAKS(r:Record, s:Worksheet):void
    {
    }

    private function HYPERLINK(r:Record, s:Worksheet):void
    {
    }

    private function INDEX(r:Record, s:Worksheet):void
    {
    }

    private function INTEGER(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();

      // Cell attributes
      var attr1:uint = r.data.readUnsignedByte();
      var attr2:uint = r.data.readUnsignedByte();
      var attr3:uint = r.data.readUnsignedByte();

      // Figure out the format
      var format:uint = attr2 & 0x3F;

      // Integer values can only be unsigned
      var value:Number = r.data.readUnsignedShort();

      // Figure out the format
      var formatString:String = s.formats[format];
      s.setCell(row, col, value);
      s.cell(row, col).format = formatString;
    }

    private function ITERATION(r:Record, s:Worksheet):void
    {
      var iteration:Boolean = r.data.readUnsignedShort() == 1;
    }

    private function IXFE(r:Record, s:Worksheet):void
    {
    }

    private function LABEL(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();

      var len:uint;

      if (r.type == Type.LABEL)
      {
        // BIFF2
        // Cell attributes
        r.data.readUnsignedByte();
        r.data.readUnsignedByte();
        r.data.readUnsignedByte();

        len = r.data.readUnsignedByte();
      }
      else
      {
        // BIFF3+
        var indexToXF:uint = r.data.readUnsignedShort();
        len = r.data.readUnsignedShort();
      }

      var value:String = r.data.readMultiByte(len, _charset);
      var fmt:String = s.formats[s.xformats[indexToXF].format];
      if (fmt == null || fmt.length == 0)
      {
        fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
      }
      s.setCell(row, col, value);
      s.cell(row, col).format = fmt;
    }

    private function LABELRANGES(r:Record, s:Worksheet):void
    {
    }

    private function LABELSST(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();

      var xfIndex:uint = r.data.readUnsignedShort();
      var sstIndex:uint = r.data.readUnsignedInt();

      var value:String = _sst[sstIndex];
      s.setCell(row, col, value);
    }

    private function LEFTMARGIN(r:Record, s:Worksheet):void
    {
      var size:Number = r.data.readDouble();
    }

    private function MERGEDCELLS(r:Record, s:Worksheet):void
    {
    }

    private function MULBLANK(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();
      while (r.data.bytesAvailable > 2)
      {
        var indexToXF:uint = r.data.readUnsignedShort();
        s.setCell(row, col, "");
        col++;
      }
    }

    private function MULRK(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();
      while (r.data.bytesAvailable > 2)
      {
        var indexToXF:uint = r.data.readUnsignedShort();
        var value:Number = readRK(r, s);
        var fmt:String = s.formats[s.xformats[indexToXF].format];
        if (fmt == null || fmt.length == 0)
        {
          fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
        }
        s.setCell(row, col, value);
        s.cell(row, col).format = fmt;

        col++;
      }
    }

    private function NOTE(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();

      var note:String;
      if (_version <= BIFFVersion.BIFF5)
      {
        var totalLength:uint = r.data.readUnsignedShort();
        note = r.data.readUTFBytes(r.data.bytesAvailable);

      }
      else
      {
        var flags:uint = r.data.readUnsignedShort();
        var idx:uint = (r.data.readUnsignedShort() - 1)*2;
        // var author:String = r.readUnicodeStr16();
        var author:String = readString(r.data);
        note = notes[idx];
      }
      s.cell(row, col).note = note;
    }

    private function NUMBER(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();

      if (r.type < 0x200)
      {
        // BIFF2
        // Cell attributes
        r.data.readUnsignedByte();
        r.data.readUnsignedByte();
        r.data.readUnsignedByte();
      }
      else
      {
        // BIFF>2
        var indexToXF:uint = r.data.readUnsignedShort();
      }


      var value:Number = r.data.readDouble();
      if (_version == BIFFVersion.BIFF2)
      {
        s.setCell(row, col, value);
      }
      else
      {
        s.setCell(row, col, value);
        var fmt:String = s.formats[s.xformats[indexToXF].format];
        if (fmt == null || fmt.length == 0)
        {
          fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
        }
        s.cell(row, col).format = fmt;
      }
    }

    private function OBJECTPROTECT(r:Record, s:Worksheet):void
    {
    }

    private function PAGESETUP(r:Record, s:Worksheet):void
    {
      var paperSize:uint = r.data.readUnsignedShort();
      var scaleFactor:uint = r.data.readUnsignedShort();
      var startPageNumber:uint = r.data.readUnsignedShort();
      var maxWidthInPages:uint = r.data.readUnsignedShort();
      var maxHeightInPages:uint = r.data.readUnsignedShort();

      var optionFlags:uint;
      if (_version == BIFFVersion.BIFF4)
      {
        optionFlags = r.data.readUnsignedShort();
      }
      else
      {
        optionFlags = r.data.readUnsignedShort();
        var printDPI:uint = r.data.readUnsignedShort();
        var verticalPrintDPI:uint = r.data.readUnsignedShort();
        var headerMargin:Number = r.data.readDouble();
        var footerMargin:Number = r.data.readDouble();
        var copies:uint = r.data.readUnsignedShort();
      }
    }

    private function PALETTE(r:Record, s:Worksheet):void
    {
    }

    private function PANE(r:Record, s:Worksheet):void
    {
    }

    private function PASSWORD(r:Record, s:Worksheet):void
    {
    }

    private function PHONETICPR(r:Record, s:Worksheet):void
    {
    }

    private function PLS(r:Record, s:Worksheet):void
    {
    }

    private function PRECISION(r:Record, s:Worksheet):void
    {
      var fullPrecision:Boolean = r.data.readUnsignedShort() == 1;
    }

    private function PRINTGRIDLINES(r:Record, s:Worksheet):void
    {
      var printGridlines:Boolean = r.data.readUnsignedShort() == 1
    }

    private function PRINTHEADERS(r:Record, s:Worksheet):void
    {
      var printHeaders:Boolean = r.data.readUnsignedShort() == 1;
    }

    private function PROTECT(r:Record, s:Worksheet):void
    {
    }

    private function QUICKTIP(r:Record, s:Worksheet):void
    {
    }

    private function RANGEPROTECTION(r:Record, s:Worksheet):void
    {
    }

    private function REFMODE(r:Record, s:Worksheet):void
    {
      var mode:uint = r.data.readUnsignedShort();
    }

    private function RIGHTMARGIN(r:Record, s:Worksheet):void
    {
      var size:Number = r.data.readDouble();
    }

    private function RK(r:Record, s:Worksheet):void
    {
      var row:uint = r.data.readUnsignedShort();
      var col:uint = r.data.readUnsignedShort();
      var indexToXF:uint = r.data.readUnsignedShort();


      var value:Number = readRK(r, s);

      s.setCell(row, col, value);

      var fmt:String = s.formats[s.xformats[indexToXF].format];
      if (fmt == null || fmt.length == 0)
      {
        fmt = Formatter.builtInFormats[s.xformats[indexToXF].format];
      }

      s.cell(row, col).format = fmt;
    }

    private function ROW(r:Record, s:Worksheet):void
    {
      var rowNum:uint = r.data.readUnsignedShort();
      var firstCol:uint = r.data.readUnsignedShort();
      var lastCol:uint = r.data.readUnsignedShort()-1;
      var rowHeight:uint = r.data.readUnsignedShort();
      var microsoftUse:uint = r.data.readUnsignedShort();
      var defaultCellAttrs:Boolean = r.data.readUnsignedByte() == 1;
      var cellRecordsOffset:uint = r.data.readUnsignedShort();
      // Last three bytes are default cell attributes
    }

    private function RSTRING(r:Record, s:Worksheet):void
    {
      LABEL(r, s);
    }

    private function SAVERECALC(r:Record, s:Worksheet):void
    {
      var recalculateBeforeSave:Boolean = r.data.readUnsignedByte() == 1;
    }

    private function SCENPROTECT(r:Record, s:Worksheet):void
    {
    }

    private function SCL(r:Record, s:Worksheet):void
    {
    }

    private function SELECTION(r:Record, s:Worksheet):void
    {
    }

    private function SHAREDFMLA(r:Record, s:Worksheet):void
    {
      var firstRow:uint = r.data.readUnsignedShort();
      var lastRow:uint = r.data.readUnsignedShort();
      var firstCol:uint = r.data.readUnsignedByte();
      var lastCol:uint = r.data.readUnsignedByte();

      r.data.position++;
      var numExistingFormulaRecords:uint = r.data.readUnsignedByte();
      var tokLen:uint = _version == BIFFVersion.BIFF2 ? r.data.readUnsignedByte() : r.data.readUnsignedShort();

      // Next comes the formula
      var tokens:ByteArray = new ByteArray();
      tokens.endian = Endian.LITTLE_ENDIAN;
      r.data.readBytes(tokens, 0, tokLen);
      // Util.debugByteArray(tokens, "SHAREDFMLA");

      for (var rw:uint = firstRow; rw < lastRow; rw++)
      {
        for (var c:uint = firstCol; c <= lastCol; c++)
        {
          s.cell(rw, c).sharedTokens = tokens;
        }
      }
    }

    private function SHEET(r:Record, s:Worksheet):void
    {
      var sheetOffset:uint = r.data.readUnsignedInt();
      var visibility:uint = r.data.readUnsignedByte();
      var sheetType:uint = r.data.readUnsignedByte();

      var len:uint, f:uint = 0;
      var name:String;
      if (_version <= BIFFVersion.BIFF5)
      {
        len = r.data.readUnsignedByte();
        name = r.data.readMultiByte(len, _charset);
      }
      else
      {
        len = r.data.readUnsignedByte();
        f = r.data.readUnsignedByte(); // ?单字节=0, 双字节=1
        name = f ? r.data.readMultiByte(2*len, _charset) : r.data.readUTFBytes(len);
      }
      var currentSheet:Worksheet;
      currentSheet = new Worksheet(this);
      currentSheet.dateMode = dateMode;
      currentSheet.name = name;
      currentSheet.formats = currentSheet.formats.concat(globalFormats);
      currentSheet.xformats = currentSheet.xformats.concat(globalXFormats);
      _sheets.addItem(currentSheet);
    }

    private function SHEETLAYOUT(r:Record, s:Worksheet):void
    {
    }

    private function SHEETPR(r:Record, s:Worksheet):void
    {
    }

    private function SHEETPROTECTION(r:Record, s:Worksheet):void
    {
    }

    private function SORT(r:Record, s:Worksheet):void
    {
    }

    private function SST(r:Record, s:Worksheet):void
    {
      numWorkbookStrings = r.data.readUnsignedInt();
      _sstCount = r.data.readUnsignedInt();
      SST2(r, s);
    }

    private function SST2(r:Record, s:Worksheet):void
    {
      var str:String, p:uint;
      rrest.data.clear();
      while (r.data.bytesAvailable > 0)
      {
        p = r.data.position;
        str = readString(r.data);
        if (rrest.e1 || rrest.e3 || rrest.e4)
        {
          rrest.data.writeBytes(r.data, p);
          break;
        }
        _sst.push(str);
      }
    }

    private function STANDARDWIDTH(r:Record, s:Worksheet):void
    {
    }

    private function STRING(r:Record, s:Worksheet):void
    {
    }

    private function STYLE(r:Record, s:Worksheet):void
    {
    }

    private function TOPMARGIN(r:Record, s:Worksheet):void
    {
      var size:Number = r.data.readDouble();
    }

    private function UNCALCED(r:Record, s:Worksheet):void
    {
      // Indicates that formulas were not recalculated before the sheet was saved
    }

    private function USESELFS(r:Record, s:Worksheet):void
    {
    }

    private function VCENTER(r:Record, s:Worksheet):void
    {
      // 0 = top align, 1 = centered
      var center:uint = r.data.readUnsignedShort();
    }

    private function VERTICALPAGEBREAKS(r:Record, s:Worksheet):void
    {
    }

    private function WINDOW1(r:Record, s:Worksheet):void
    {
      var x:uint = r.data.readUnsignedShort();
      var y:uint = r.data.readUnsignedShort();
      var width:uint = r.data.readUnsignedShort();
      var height:uint = r.data.readUnsignedShort();
      var hidden:Boolean = r.data.readUnsignedByte() == 1;
    }

    private function WINDOW2(r:Record, s:Worksheet):void
    {
      if (r.type == 0x003E)
      {
        // BIFF2
        var showFormulas:Boolean = r.data.readUnsignedByte() == 1;
        var showGridLines:Boolean = r.data.readUnsignedByte() == 1;
        var showRowColHeaders:Boolean = r.data.readUnsignedByte() == 1;
        var frozen:Boolean = r.data.readUnsignedByte() == 1;
        var showZeros:Boolean = r.data.readUnsignedByte() == 1;
        var topRowVisible:uint = r.data.readUnsignedShort();
        var leftColVisible:uint = r.data.readUnsignedShort();
        var showHeadersDefaultColor:Boolean = r.data.readUnsignedByte() == 1;
        var headerColor:uint = r.data.readUnsignedInt();
      }
      else
      {
        // BIFF>2
        var options:uint = r.data.readUnsignedShort();
        leftColVisible = r.data.readUnsignedShort();
        showHeadersDefaultColor = r.data.readUnsignedByte() == 1;
        headerColor = r.data.readUnsignedInt();
      }
    }

    private function WINDOWPROTECT(r:Record, s:Worksheet):void
    {
      var windowsProtected:Boolean = r.data.readUnsignedShort() == 1;
    }

    private function WRITEACCESS(r:Record, s:Worksheet):void
    {
      var len:uint = r.data.readUnsignedByte();
      var username:String = r.data.readUTFBytes(len);
    }

    private function WRITEPROT(r:Record, s:Worksheet):void
    {
    }

    private function XCT(r:Record, s:Worksheet):void
    {
    }

    private function XF(r:Record, s:Worksheet):void
    {
      var font:uint;
      var format:uint;
      switch (_version)
      {
        case BIFFVersion.BIFF2:
          font = r.data.readUnsignedByte();
          r.data.position++;
          format = r.data.readUnsignedByte() & 0x3F;
          break;
        case BIFFVersion.BIFF3:
        case BIFFVersion.BIFF4:
          font = r.data.readUnsignedByte();
          format = r.data.readUnsignedByte();
          break;
        case BIFFVersion.BIFF5:
          font = r.data.readUnsignedShort();
          format = r.data.readUnsignedShort();
          break;
        case BIFFVersion.BIFF8:
          font = r.data.readUnsignedShort();
          format = r.data.readUnsignedShort();
          break;
      }

      if (s is Worksheet)
      {
        s.xformats.push(new XFormat(r.type, format));
      }
      else
      {
        globalXFormats.push(new XFormat(r.type, format));
      }
    }

  }
}