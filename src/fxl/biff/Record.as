package fxl.biff
{
  import flash.utils.ByteArray;
  import flash.utils.Endian;

  /**
   * Represents a single BIFF record.
   */
  public class Record
  {
    private var _type:uint;
    private var _data:ByteArray;

    /**
     * @param type The type of the BIFF record
     * @param data A byte array containing the BIFF record without the length and type header
     */
    public function Record(type:uint, data:ByteArray = null)
    {
      _type = type;
      _data = data == null ? new ByteArray() : data;
      _data.endian = Endian.LITTLE_ENDIAN;
    }

    public function get type():uint
    {
      return _type;
    }

    public function set type(value:uint):void
    {
      _type = value;
    }

    public function get data():ByteArray
    {
      return _data;
    }

    public function get length():uint
    {
      return _data.length;
    }

    public function insert(b:ByteArray, pos:int=0):void
    {
      if (b == null || b.length <= 0)
        return;
      if (pos < 0)
        pos = b.length;
      var _d:ByteArray = new ByteArray();
      if (pos > 0)
        _d.writeBytes(_data, 0, pos);
      _d.writeBytes(b, 0, b.length);
      if (_data.length > pos)
        _d.writeBytes(_data, pos, _data.length-pos);
      _d.endian = Endian.LITTLE_ENDIAN;
      _d.position = 0;
      _data = _d;
    }

    public function append(b:ByteArray):void
    {
      insert(b, -1);
    }

    // CONTINUE的SST: 3C 00 47 06 01, 需要把第5个字节的00或01去掉
    public function ltrim():void
    {
      if (_data == null || _data.length <= 0)
        return;
      var p:uint = _data.position;
      var i:uint, b:uint;
      var _d:ByteArray = new ByteArray();
      _data.position = 0;
      b = _data.readUnsignedByte();
      if ((b == 0x00 || b == 0x01) && _data.position < _data.length)
      {
        _d.writeBytes(_data, _data.position);
        _d.position = p;
        _d.endian = Endian.LITTLE_ENDIAN;
        _data = _d;
      }
      _data.position = p;
    }

    /**
     * @deprecated
     * 	Reads one of the weird Excel 2 byte length unicode strings.
     *  @return A unicode string with 2 byte length read from the _data ByteArray's current position.
     */
    protected function readUnicodeStr16():String
    {
      var len:uint = _data.readUnsignedShort();
      var opts:uint = _data.readByte();
      var compressed:Boolean = (opts & 0x01) == 0;
      var asianPhonetic:Boolean = (opts & 0x04) == 0x04;
      var richtext:Boolean = (opts & 0x08) == 0x08;

      if (!compressed)
      {
        len *= 2;
      }

      len = len > _data.bytesAvailable ? _data.bytesAvailable : len;
      var ret:String = _data.readUTFBytes(len);
      return ret;
    }
  }
}