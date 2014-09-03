package fxl
{
  import flash.utils.ByteArray;

  public class Util
  {
    public function Util()
    {
    }

    /**
     * 输出ByteArray的十六进制值
     **/
    public static function debugByteArray(bytes:ByteArray, tip:String="", count:uint=2):void
    {
      var s:String="", h:String;
      var b:uint;
      var i:int=0, p:int=bytes.position;
      trace(tip + "[position/length]: " + p + "/" + bytes.length); // 53562
      bytes.position=0;
      while (bytes.bytesAvailable > 0)
      {
        i++;
        b=bytes.readByte();
        h=b.toString(16);
        h=(h.length < 2) ? "0" + h : h.substr(h.length-2, 2);
        s=s + h.toUpperCase() + " ";
        if (i%16==0)
          s=s + "\n";
        if (count > 0 && i==16*count)
        {
          if (bytes.length > 32*count)
          {
            s=s + "... ...\n";
            bytes.position=int((bytes.length-1)/16-count+1) * 16;
          }
          count--;
        }
      }
      trace(s);
      bytes.position=p;
    }
  }
}