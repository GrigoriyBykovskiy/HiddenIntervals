using System;
using System.Linq;
using System.Text;

namespace HiddenInvervals
{
    public class SteganoModel
    {
        public static string Encrypt(string containerData, string messageData)
        {
            var spaceCount = containerData.Count(x => x == ' ');

            var buf = new StringBuilder();
            foreach (byte b in System.Text.Encoding.UTF8.GetBytes(messageData))
                buf.Append(  Convert.ToString(b, 2).PadLeft(8,'0'));
 
            var messageBinaryData = buf.ToString();

            if (spaceCount < messageBinaryData.Length)
                throw new Exception("Anta Baka?! Сообщение слишком большое, сэмпай.");

            var indexCounter = 0;

            var result = new StringBuilder();

            foreach (var ch in containerData)
            {
                result.Append(ch);
                if (ch == ' ' && indexCounter < messageBinaryData.Length)
                {
                    if (messageBinaryData[indexCounter] == '1')
                    {
                        result.Append(' ');
                    }
                    indexCounter++;
                }
            }
            return result.ToString();
        }

        public static string Decrypt(string containerData)
        {
            var buf = new StringBuilder();
            
            for (var i = 0; i < containerData.Length; i++)
            {
                if (containerData[i] != ' ') continue;
                if (i + 1 < containerData.Length)
                {
                    if (containerData[i + 1] == ' ')
                    {
                        buf.Append('1');
                        i++;
                    }
                    else
                    {
                        buf.Append('0');
                    }
                }
            }
            
            while (buf.Length % 8 != 0)
            {
                buf.Append('0');
            }
            
            var messageBinary = buf.ToString();

            var arr = new byte[messageBinary.Length / 8];
            var count = 0;
            
            for (var i = 0; i < messageBinary.Length; i += 8)
            {
                var ss = messageBinary.Substring(count * 8, 8);
                var bv = Convert.ToByte(ss, 2);
                arr[i / 8] = bv;
                count++;
            }

            return Encoding.UTF8.GetString(arr);
        }
    }
}