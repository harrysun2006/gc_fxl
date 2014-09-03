package fxl.biff
{
  import flash.utils.ByteArray;
  import flash.utils.Endian;

  /**
   * Used to make reading BIFF files a little less painful.
   */
  public class BIFFReader
  {
    private var stream:ByteArray;
    private var _tags:Array;


    /**
     *
     * @param stream A ByteArray containing a BIFF document as a stream. The ByteArray will be rewound and set to LITTLE_ENDIAN
     * automatically.
     *
     */
    public function BIFFReader(stream:ByteArray)
    {
      this.stream = stream;
      stream.endian = Endian.LITTLE_ENDIAN;
      stream.position = 0;
    }

    public function get position():uint
    {
      return stream.position;
    }

    public function get available():uint
    {
      return stream.bytesAvailable;
    }

    public function get length():uint
    {
      return stream.length;
    }

    /**
     * Fetches the next Record in the BIFF file or null if there are no records left
     *
     * @return The next Record in the file or null if no more Records are available
     *
     * @see com.as3xls.biff.Record
     */
    public function readTag():Record
    {
      if (stream.bytesAvailable < 4)
      {
        return null;
      }

      var type:uint = stream.readUnsignedShort();
      var length:uint = stream.readUnsignedShort();

      if (stream.bytesAvailable < length)
      {
        return null;
      }

      var ret:ByteArray = new ByteArray();
      ret.endian = Endian.LITTLE_ENDIAN;

      // Length of 0 will actually read the whole stream, which is bad
      if (length > 0)
      {
        stream.readBytes(ret, 0, length);
      }
      return new Record(type, ret);
    }

  }
}