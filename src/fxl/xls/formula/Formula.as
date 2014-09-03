package fxl.xls.formula
{
  import fxl.biff.BIFFVersion;
  import fxl.xls.Cell;
  import fxl.xls.Workbook;
  import fxl.xls.Worksheet;

  import flash.utils.ByteArray;
  import flash.utils.Endian;

  /**
   * <p>
   * Represents Formulas in Excel formulas and handles converting them to human readable format and evaluating them.
   * </p>
   *
   * <p>
   * Excel stores formulas as RPN tokens. This makes it relatively easy to evaluate them but a bit of a pain to display
   * them in human readable format. You can't please everyone, I guess.
   * </p>
   *
   */
  public class Formula
  {
    private var _book:Workbook;
    private var _sheet:Worksheet;
    private var _tokens:ByteArray;
    private var _row:uint;
    private var _col:uint;
    private var _result:*;
    private var _result0:Number;
    private var _biff:uint;
    private var _alwaysRecalculate:Boolean;
    private var _formula:String;

    private var dirty:Boolean;

    private var isSharedFormula:Boolean;
    private var sharedFormulaRow:uint;
    private var sharedFormulaCol:uint;
    private static const MISSING:Object = {id:"M"};

    /**
     * Creates a new Formula in a given cell with an array of tokens. Note that the row, col, and so on don't
     * affect the actual location of the formula. They are only used when relative cell locations rear their
     * ugly but admittedly useful heads. The location is determined by this formula's cell object
     *
     * @param book: The book which the formula inhabits, useful for formulas of global defined names
     * @param sheet: The sheet which the formula inhabits
     * @param row: Formula's cell's row
     * @param col: Formula's cell's col
     * @param tokens: A ByteArray containing the tokens in Excel's native RPN format
     * @result0: Result(Number) saved redundantly in formula
     *
     * @see fxl.xls.Cell
     *
     */
    public function Formula(book:Workbook, sheet:Worksheet, row:uint, col:uint, tokens:ByteArray, result0:Number=NaN)
    {
      _book = book;
      _sheet = sheet;
      _row = row;
      _col = col;
      _result0 = result0;
      _result = isNaN(_result0) ? result : _result0;
      _tokens = tokens;
      _tokens.endian = flash.utils.Endian.LITTLE_ENDIAN;
      _biff = _book ? _book.version : BIFFVersion.BIFF0;
      _formula = "";
      isSharedFormula = false;
      dirty = true;
    }

    /**
     * Sums the Numbers contained in the given array
     * @param args An array of Numbers to add up
     * @return The sum of the given numbers
     *
     */
    private function builtInSum(args:Array):Number
    {
      var ret:Number = 0;
      for (var n:uint = 0; n < args.length; n++)
      {
        var num:Number = Number(args[n]);
        ret += num;
      }
      return ret;
    }

    /**
     * Recursively converts the given object into a one-dimensional array. If an array of multiple dimensions is passed in
     * the result will still be a one-dimensional array containing each element.
     *
     * @param arg The object to convert
     * @return A one-dimensional array containing the element passed in or elements of the array passed in.
     *
     */
    private function convertToArray(arg:*):Array
    {
      var ret:Array = new Array();
      if (arg is Array)
      {
        for (var i:uint = 0; i < arg.length; i++)
        {
          ret = ret.concat(convertToArray(arg[i]));
        }
      }
      else
      {
        ret = [arg];
      }
      return ret;
    }

    private function builtInSubstitute(v1:Object, v2:Object, v3:Object, v4:Object=null):String
    {
      var s1:String = String(v1);
      var s2:String = String(v2);
      var s3:String = v3 == MISSING ? "" : String(v3);
      var i:uint = v4 ? Number(v4) : 0;
      var r:String = "";
      var i1:int, i2:int, i3:int;
      i1 = i2 = i3 = 0;
      if (s2 == null || s2 == "")
        return s1;
      while (i1 < s1.length)
      {
        i2 = s1.indexOf(s2, i1);
        if (i2 < 0)
        {
          r += s1.substring(i1);
          break;
        }
        else
        {
          i3++;
          if (i == 0 || i == i3)
          {
            r += s1.substring(i1, i2)+s3;
          }
          i1 = i2+s2.length;
        }
      }
      return r;
    }

    private function builtInVLookup(s:String, cells:Array, col:Number, like:Boolean):Object
    {
      var r:uint, c:uint;
      var rr:Array;
      for (r = 0; r < cells.length; r++)
      {
        rr = cells[r] as Array;
        if ((like && rr[0].indexOf(s) >= 0) || rr[0] == s)
          return rr[col-1];
      }
      return "";
    }

    /**
     * Executes one of Excel's built in functions (obviously without relying on Excel).
     *
     * @param idx The idx of the function to use
     * @param rest The arguments of the function to be called
     * @return The result of the operation
     *
     */
    private function builtInFunction(idx:uint, ... rest):*
    {
      var ret:*;
      var a:Array;
      var n:Number;
      switch (idx)
      {
        case 0: // COUNT
          return convertToArray(rest).length;
        case 1: // IF
          return Boolean(rest[0]) ? rest[1] : rest[2];
        case 4: // SUM
          ret = 0;
          a = convertToArray(rest);
          for (var i:uint = 0; i < a.length; i++)
          {
            ret += a[i];
          }
          return ret;
        case 5: // AVERAGE
          a = convertToArray(rest);
          return builtInSum(a) / a.length;
        case 6: // MIN
          return Math.min.apply(null, convertToArray(rest));
        case 7: // MAX
          return Math.max.apply(null, convertToArray(rest));
        case 15: // SIN
          return Math.sin(rest[0]);
        case 16: // COS
          return Math.cos(rest[0]);
        case 17: // TAN
          return Math.tan(rest[0]);
        case 18: // ARCTAN
          return Math.atan(rest[0]);
        case 19: // PI
          return Math.PI;
        case 20: // SQRT
          return Math.sqrt(rest[0]);
        case 21: // EXP
          return Math.exp(rest[0]);
        case 22: // LN
          return Math.log(rest[0]) / Math.LOG10E;
        case 23: // LOG10	
          return Math.log(rest[0]);
        case 24: // ABS
          return Math.abs(rest[0]);
        case 25: // INT
          return Math.round(rest[0]);
        case 30: // REPT
          var text:String = String(rest[0]);
          var count:Number = Number(rest[1]);
          ret = "";
          for (n = 0; n < count; n++)
          {
            ret += text;
          }
          return ret;
        case 31: // MID
          return String(rest[0]).substr(Number(rest[1]), Number(rest[2]));
        case 32: // LEN
          return String(rest[0]).length;
        case 34: // TRUE
          return true;
        case 35: // FALSE
          return false;
        case 36: // AND
          for (n = 0; n < rest.length; n++)
          {
            if (Boolean(rest[n]) == false)
            {
              return false;
            }
          }
          return true;
        case 37: // OR
          for (n = 0; n < rest.length; n++)
          {
            if (Boolean(rest[n]) == true)
            {
              return true;
            }
          }
          return false;
        case 38: // NOT
          return !Boolean(rest[0]);
        case 39: // MOD
          return Number(rest[0]) % Number(rest[1]);
        case 56: // PV
          var rate:Number = Number(rest[0]);
          var nper:Number = Number(rest[1]);
          var pmt:Number = Number(rest[2]);
          return -pmt*((Math.pow(1+rate, nper)-1)/rate) / (Math.pow(1+rate, nper));
        case 63: // RAND
          return Math.random();
        case 97: // ATAN2
          return Math.atan2(Number(rest[0]), Number(rest[1]));
        case 98: // ASIN
          return Math.asin(Number(rest[0]));
        case 99: // ACOS
          return Math.acos(Number(rest[0]));
        case 102: // VLOOKUP
          return builtInVLookup(String(rest[0]), rest[1], Number(rest[2]), Boolean(rest[3]));
        case 109: // LOG
          return Math.log(rest[0]);
        case 111: // CHAR
          return String.fromCharCode(rest[0]);
        case 112: // LOWER
          return String(rest[0]).toLowerCase();
        case 113: // UPPER
          return String(rest[0]).toUpperCase();
        case 115: // LEFT
          return String(rest[0]).substr(0, Number(rest[1]));
        case 116: // RIGHT
          return String(rest[0]).substr(String(rest[0]).length-Number(rest[1]), Number(rest[1]));
        case 118: // TRIM
          return String(rest[0]).replace(/(^\s*)|(\s*$)/g,"");
        case 120: // SUBSTITUTE
          return (rest.length >= 3) ? builtInSubstitute(rest[0], rest[1], rest[2], rest.length > 3 ? rest[3] : null) : "";
        default:
          throw new Error("Unsupported function: " + idx);
      }
    }

    /**
     * Evaluates the represented function and stores the result
     *
     */
    public function updateResult():void
    {
      if (_tokens == null)
      {
        return;
      }
      _tokens.position = 0;

      var tok:uint;
      var v1:*;
      var v2:*;
      var idx:uint;
      var n:uint;
      var numArgs:uint;
      var args:Array;
      var rsi:uint; // ref-sheet index
      var rs:Worksheet; // ref-sheet
      var r1:uint;
      var c1:uint;
      var r2:uint;
      var c2:uint;
      var cell:Cell;
      var rni:uint; // ref-name index
      var rn:Object; // ref-name
      var cells:Array;
      var r:uint, c:uint;

      var stack:Array = new Array();
      var unknown:Array = new Array();

      while (_tokens.bytesAvailable > 0)
      {
        tok = _tokens.readUnsignedByte();

        switch (tok)
        {
          case Tokens.tExp:
            r1 = _tokens.readUnsignedShort();
            c1 = _biff == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
            sharedFormulaRow = r1;
            sharedFormulaCol = c1;

            if (_sheet.cell(r1, c1).sharedTokens == null)
            {
              return;
            }
            _tokens = _sheet.cell(r1, c1).sharedTokens;
            updateResult();
            return;
            break;
          // Constant tokens
          case Tokens.tBool:
            stack.push(_tokens.readUnsignedByte() == 1);
            break;
          case Tokens.tNum:
            stack.push(_tokens.readDouble());
            break;
          case Tokens.tInt:
            stack.push(_tokens.readUnsignedShort());
            break;
          case Tokens.tStr:
            var len:uint = _tokens.readUnsignedByte();
            if (_biff == BIFFVersion.BIFF8)
            {
              _tokens.position++; // Skip option byte
            }

            stack.push(_tokens.readUTFBytes(len));
            break;

          // Unary oporators
          case Tokens.tUplus:
            stack.push(stack.pop());
            break;
          case Tokens.tUminus:
            stack.push(stack.pop() * -1);
            break;
          case Tokens.tPercent:
            stack.push(stack.pop() / 100);
            break;

          // Binary operators
          case Tokens.tAdd:
            stack.push(stack.pop() + stack.pop());
            break;
          case Tokens.tSub:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 - v1);
            break;
          case Tokens.tMul:
            stack.push(stack.pop() * stack.pop());
            break;
          case Tokens.tDiv:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 / v1);
            break;
          case Tokens.tPower:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(Math.pow(v2, v1));
            break;
          case Tokens.tLT:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 < v1);
            break;
          case Tokens.tLE:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 <= v1);
            break;
          case Tokens.tEQ:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 == v1);
            break;
          case Tokens.tGE:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 >= v1);
            break;
          case Tokens.tGT:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 > v1);
            break;
          case Tokens.tNE:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 != v1);
            break;
          case Tokens.tConcat:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + v1);
            break;
          case Tokens.tParen:
            break;

          // Functions
          case Tokens.tFuncR:
          case Tokens.tFuncV:
          case Tokens.tFuncA:
            idx = _biff <= BIFFVersion.BIFF3 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
            numArgs = Functions.numArgs[idx];
            args = new Array();
            for (n = 0; n < numArgs; n++)
            {
              args.push(stack.pop());
            }
            args.push(idx);
            stack.push(builtInFunction.apply(this, args.reverse()));
            break;
          case Tokens.tFuncVarR:
          case Tokens.tFuncVarV:
          case Tokens.tFuncVarA:
            if (_biff <= BIFFVersion.BIFF3)
            {
              numArgs = _tokens.readUnsignedByte();
              idx = _tokens.readUnsignedByte();
            }
            else
            {
              numArgs = _tokens.readUnsignedByte() & 0x7F;
              idx = _tokens.readUnsignedShort() & 0x7FFF;
            }
            args = new Array();
            for (n = 0; n < numArgs; n++)
            {
              args.push(stack.pop());
            }
            args.push(idx);
            stack.push(builtInFunction.apply(this, args.reverse()));
            break;
          case Tokens.tFuncCER:
          case Tokens.tFuncCEV:
          case Tokens.tFuncCEA:
            numArgs = _tokens.readUnsignedByte();
            idx = _tokens.readUnsignedByte();
            args = new Array();
            for (n = 0; n < numArgs; n++)
            {
              args.push(stack.pop());
            }
            args.push(idx);
            stack.push(builtInFunction.apply(this, args.reverse()));
            break;
          case Tokens.tRefNR:
          case Tokens.tRefNV:
          case Tokens.tRefNA:
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readShort();
              c1 = _tokens.readByte();
              var flag:uint = _tokens.readUnsignedByte();

              if ((flag & 0x80) != 0)
              {
                r1 = _row + r1;
              }
              if ((flag & 0x40) != 0)
              {
                c1 = _col + c1;
              }
            }
            else
            {
              r1 = _tokens.readShort() + 1;
              c1 = _tokens.readByte();
              if (r1 & 0x8000 != 0)
              {
                c1 = _col + c1;
              }
              if (r1 & 0x4000 != 0)
              {
                r1 = _row + r1;
              }
              r1 &= 0x3FFF;
            }

            // HACK Seems BIFF5 is off by one???
            if (_biff == BIFFVersion.BIFF5)
            {
              r1--;
            }

            stack.push(_sheet.cell(r1, c1).value);
            break;
          case Tokens.tRefR:
          case Tokens.tRefV:
          case Tokens.tRefA:
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readUnsignedShort();
              c1 = _tokens.readUnsignedShort() & 0x00FF;
            }
            else
            {
              r1 = _tokens.readShort() & 0x3FFF;
              c1 = _tokens.readByte();
            }

            stack.push(_sheet.cell(r1, c1).value);
            break;
          case Tokens.tAreaR:
          case Tokens.tAreaV:
          case Tokens.tAreaA:
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readUnsignedShort();
              r2 = _tokens.readUnsignedShort();
              c1 = _tokens.readUnsignedShort() & 0x00FF;
              c2 = _tokens.readUnsignedShort() & 0x00FF;
            }
            else
            {
              r1 = _tokens.readUnsignedShort() & 0x3FFF;
              r2 = _tokens.readUnsignedShort() & 0x3FFF;
              c1 = _tokens.readUnsignedByte();
              c2 = _tokens.readUnsignedByte();
            }

            cells = new Array();
            for (r = r1; r <= r2; r++)
            {
              cells[r-r1] = new Array();
              for (c = c1; c <= c2; c++)
              {
                cells[r-r1][c-c1] = _sheet.cell(r, c).value;
              }
            }
            stack.push(cells);
            break;
          case Tokens.tAttr:
            var attrType:uint = _tokens.readUnsignedByte();
            var dist:uint;
            switch (attrType)
          {
            case 0x02: // tAttrIF: IF Token control
              dist = _biff == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
              if (stack.pop() == false)
              {
                _tokens.position += dist;
              }
              break;
            case 0x08: // tAttrSkip: Jump-like command
              dist = _biff == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte()+1 : _tokens.readUnsignedShort()+1;
              _tokens.position += dist;
              break;
            case 0x10: // tAttrSum: SUM with 1 parameter
              _tokens.position += _biff == BIFFVersion.BIFF2 ? 1 : 2;
              stack.push(builtInSum(convertToArray(stack.pop())));
              break;
            case 0x40: // tAttrSpace: Need to skip an extra space
              _tokens.position += _biff == BIFFVersion.BIFF2 ? 1 : 2;
              break;
          }
            break;
          case Tokens.NOTUSED:
            break;
          case Tokens.tRef3dR:
          case Tokens.tRef3dV:
          case Tokens.tRef3dA:
            // CommonUtil.debugByteArray(_tokens, false);
            rsi = _tokens.readUnsignedShort();
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readUnsignedShort();
              c1 = _tokens.readUnsignedShort() & 0x00FF;
            }
            else
            {
              r1 = _tokens.readUnsignedShort() & 0x3FFF;
              c1 = _tokens.readUnsignedByte();
            }
            rs = _book.sheet(rsi);
            cell = rs ? rs.cell(r1, c1) : null;
            stack.push(cell ? cell.value : _result0);
            break;
          case Tokens.tNameR:
          case Tokens.tNameV:
          case Tokens.tNameA:
            rni = _tokens.readUnsignedShort();
            rn = _book.name(rni);
            var ff:Formula = rn.formula;
            stack.push(ff.result);
            _tokens.position += 2; // not used
            break;
          case Tokens.tArea3dR:
          case Tokens.tArea3dV:
          case Tokens.tArea3dA:
            rsi = _tokens.readUnsignedShort();
            rs = _book.sheet(rsi);
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readUnsignedShort();
              r2 = _tokens.readUnsignedShort();
              c1 = _tokens.readUnsignedShort() & 0x00FF;
              c2 = _tokens.readUnsignedShort() & 0x00FF;
            }
            else
            {
              r1 = _tokens.readUnsignedShort() & 0x3FFF;
              r2 = _tokens.readUnsignedShort() & 0x3FFF;
              c1 = _tokens.readUnsignedByte();
              c2 = _tokens.readUnsignedByte();
            }
            r2 = r2 > rs.rows ? rs.rows : r2;
            c2 = c2 > rs.cols ? rs.cols : c2;

            cells = new Array();
            for (r = r1; r <= r2; r++)
            {
              cells[r-r1] = new Array();
              for (c = c1; c <= c2; c++)
              {
                cells[r-r1][c-c1] = rs.cell(r, c).value;
              }
            }
            stack.push(cells);
            break;
          case Tokens.tRefErrR:
          case Tokens.tRefErrV:
          case Tokens.tRefErrA:
            _tokens.position += _biff == BIFFVersion.BIFF8 ? 4 : 3;
            stack.push("#REF!");
            break;
          case Tokens.tMissArg:
            stack.push(MISSING);
            break;
          case Tokens.tAreaNR:
          case Tokens.tAreaNV:
          case Tokens.tAreaNA:
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readUnsignedShort();
              r2 = _tokens.readUnsignedShort();
              c1 = _tokens.readUnsignedShort() & 0x00FF;
              c2 = _tokens.readUnsignedShort() & 0x00FF;
            }
            else
            {
              r1 = _tokens.readUnsignedShort() & 0x3FFF;
              r2 = _tokens.readUnsignedShort() & 0x3FFF;
              c1 = _tokens.readUnsignedByte();
              c2 = _tokens.readUnsignedByte();
            }

            cells = new Array();
            for (r = r1; r <= r2; r++)
            {
              cells[r-r1] = new Array();
              for (c = c1; c <= c2; c++)
              {
                cells[r-r1][c-c1] = rs.cell(r, c).value;
              }
            }
            stack.push(cells);
            break;
          default:
            // CommonUtil.debugByteArray(_tokens, false);
            unknown.push(tok);
            break;
        }
      }
      if (unknown.length > 0)
      {
        // throw new Error("Unknown formula tokens: " + unknown.join(", "));
        trace("Unknown formula tokens: " + unknown.join(", "));
      }
      _result = stack.pop();
    }

    /**
     * Updates the text representation of the formula
     */
    public function updateText():void
    {
      _tokens.position = 0;

      var tok:uint;
      var v1:*;
      var v2:*;
      var idx:uint;
      var n:uint;
      var numArgs:uint;
      var args:Array;
      var r1:int;
      var c1:int;
      var r2:int;
      var c2:int;

      var stack:Array = new Array();
      var unknown:Array = new Array();

      while (_tokens.bytesAvailable > 0)
      {
        tok = _tokens.readUnsignedByte();

        switch (tok)
        {
          case Tokens.tExp:
            r1 = _tokens.readUnsignedShort();
            c1 = _biff == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
            sharedFormulaRow = r1;
            sharedFormulaCol = c1;

            if (_sheet.cell(r1, c1).sharedTokens == null)
            {
              return;
            }
            _tokens = _sheet.cell(r1, c1).sharedTokens;
            updateText();
            return;
            break;


          // Constant Operators
          case Tokens.tBool:
            v1 = _tokens.readUnsignedByte() == 1;
            stack.push(v1.toString().toUpperCase());
            break;
          case Tokens.tNum:
            stack.push(_tokens.readDouble());
            break;
          case Tokens.tInt:
            stack.push(_tokens.readUnsignedShort());
            break;
          case Tokens.tStr:
            var len:uint = _tokens.readUnsignedByte();
            if (_biff == BIFFVersion.BIFF8)
            {
              _tokens.position++; // Skip option byte
            }

            stack.push('"' + _tokens.readUTFBytes(len) + '"');
            break;

          // Unary Operators
          case Tokens.tUplus:
            stack.push("+" + stack.pop());
            break;
          case Tokens.tUminus:
            stack.push("-" + stack.pop());
            break;
          case Tokens.tPercent:
            stack.push(stack.pop() + "%");
            break;


          // Binary Operators	
          case Tokens.tAdd:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "+" + v1);
            break;
          case Tokens.tSub:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "-" + v1);
            break;
          case Tokens.tMul:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "*" + v1);
            break;
          case Tokens.tDiv:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "/" + v1);
            break;
          case Tokens.tPower:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "^" + v1);
            break;
          case Tokens.tConcat:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push('"' + v2 + '"' + "&" + '"' + v1 + '"');
            break;


          case Tokens.tParen:
            stack.push("(" + stack.pop() + ")");
            break;


          // Logic operators
          case Tokens.tLT:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "<" + v1);
            break;
          case Tokens.tLE:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "<=" + v1);
            break;
          case Tokens.tEQ:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "=" + v1);
            break;
          case Tokens.tGE:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + ">=" + v1);
            break;
          case Tokens.tGT:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + ">" + v1);
            break;
          case Tokens.tNE:
            v1 = stack.pop();
            v2 = stack.pop();
            stack.push(v2 + "!=" + v1);
            break;

          // Function operators
          case Tokens.tFuncR:
          case Tokens.tFuncV:
          case Tokens.tFuncA:
            idx = _biff <= BIFFVersion.BIFF3 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
            args = new Array();
            numArgs = Functions.numArgs[idx];
            for (n = 0; n < numArgs; n++)
            {
              args.push(stack.pop());
            }
            stack.push(Functions.names[idx] + "(" + args.join(",") + ")");
            break;
          case Tokens.tFuncVarR:
          case Tokens.tFuncVarV:
          case Tokens.tFuncVarA:
            if (_biff <= BIFFVersion.BIFF3)
            {
              numArgs = _tokens.readUnsignedByte();
              idx = _tokens.readUnsignedByte();
            }
            else
            {
              numArgs = _tokens.readUnsignedByte() & 0x7F;
              idx = _tokens.readUnsignedShort() & 0x7FFF;
            }
            args = new Array();
            for (n = 0; n < numArgs; n++)
            {
              args.push(stack.pop());
            }
            stack.push(Functions.names[idx] + "(" + args.reverse().join(",") + ")");
            break;
          case Tokens.tFuncCER:
          case Tokens.tFuncCEV:
          case Tokens.tFuncCEA:
            numArgs = _tokens.readUnsignedByte();
            idx = _tokens.readUnsignedByte();
            args = new Array();
            for (n = 0; n < numArgs; n++)
            {
              args.push(stack.pop());
            }
            stack.push(Functions.names[idx] + "(" + args.reverse().join(",") + ")");
            break;
          case Tokens.tRefNR:
          case Tokens.tRefNV:
          case Tokens.tRefNA:
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readShort();
              c1 = _tokens.readByte();
              var flag:uint = _tokens.readUnsignedByte();

              if ((flag & 0x80) != 0)
              {
                r1 = _row + r1;
              }
              if ((flag & 0x40) != 0)
              {
                c1 = _col + c1;
              }
            }
            else
            {
              r1 = _tokens.readShort() + 1;
              c1 = _tokens.readByte();
              if (r1 & 0x8000 != 0)
              {
                c1 = _col + c1;
              }
              if (r1 & 0x4000 != 0)
              {
                r1 = _row + r1;
              }
              r1 &= 0x3FFF;
            }

            // HACK Seems BIFF5 is off by one???
            if (_biff == BIFFVersion.BIFF5)
            {
              r1--;
            }

            stack.push(String.fromCharCode(c1 + 0x41) + (r1+1));
            break;
          case Tokens.tRefR:
          case Tokens.tRefV:
          case Tokens.tRefA:
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readUnsignedShort();
              c1 = _tokens.readUnsignedShort() & 0x00FF;
            }
            else
            {
              r1 = _tokens.readShort() & 0x3FFF;
              c1 = _tokens.readByte();
            }

            stack.push(String.fromCharCode(c1 + 0x41) + (r1+1));
            break;
          case Tokens.tAreaR:
          case Tokens.tAreaV:
          case Tokens.tAreaA:
            if (_biff == BIFFVersion.BIFF8)
            {
              r1 = _tokens.readUnsignedShort();
              r2 = _tokens.readUnsignedShort();
              c1 = _tokens.readUnsignedShort() & 0x00FF;
              c2 = _tokens.readUnsignedShort() & 0x00FF;
            }
            else
            {
              r1 = _tokens.readUnsignedShort() & 0x3FFF;
              r2 = _tokens.readUnsignedShort() & 0x3FFF;
              c1 = _tokens.readUnsignedByte();
              c2 = _tokens.readUnsignedByte();
            }

            stack.push(String.fromCharCode(c1 + 0x41) + (r1+1) + ":" + String.fromCharCode(c2 + 0x41) + (r2+1));
            break;

          case Tokens.tAttr:
            var attrType:uint = _tokens.readUnsignedByte();
            var dist:uint;
            switch (attrType)
          {
            case 0x02: // tAttrIF: IF Token control
              dist = _biff == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte() : _tokens.readUnsignedShort();
              break;
            case 0x08: // tAttrSkip: Jump-like command
              dist = _biff == BIFFVersion.BIFF2 ? _tokens.readUnsignedByte()+1 : _tokens.readUnsignedShort()+1;
              break;
            case 0x10: // tAttrSum: SUM with 1 parameter
              _tokens.position += _biff == BIFFVersion.BIFF2 ? 1 : 2;
              stack.push("SUM(" + stack.pop() + ")");
              break;
            case 0x40: // tAttrSpace: Need to skip an extra space
              _tokens.position += _biff == BIFFVersion.BIFF2 ? 1 : 2;
              break;
          }
            break;
          default:
            unknown.push(tok);
            break;
        }

      }
      if (unknown.length > 0)
      {
        throw new Error("Unknown formula tokens: " + unknown.join(", "));
      }
      _formula = "=" + stack.pop();
      dirty = false;
    }

    public function get formula():String
    {
      if (dirty)
      {
        updateText();
      }
      return _formula;
    }

    public function get result():*
    {
      if (_result == undefined)
      {
        try
        {
          updateResult();
        }
        catch (e1:Error)
        {
        }
      }
      return _result;
    }

    public function set result(value:*):void
    {
      _result = value;
    }

    public function get row():uint
    {
      return _row;
    }

    public function get col():uint
    {
      return _col;
    }
  }
}