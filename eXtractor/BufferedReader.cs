using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eXtractor
{
    sealed class BufferedRader : IDisposable
    {
        private StreamReader streamReader;

        private char[] charBuffer1, charBuffer2;

        private char[] currentBuffer, standByBuffer;

        private int index;

        private bool usingBuffer1, fileEnded = false;

        private Task<int> reloadTask;

        public readonly int bufferSize;

        public int currentBufferChars;

        //private System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();

        public char GetNextChar()
        {
            if (index >= currentBufferChars) // current buffer is done. switch to next
            {
                // switch buffer
                
                //watch.Restart();
                reloadTask.Wait();
                //if (usingBuffer1)
                //    Console.WriteLine("switching to buffer 2. waited for "+watch.ElapsedMilliseconds);
                //else
                //    Console.WriteLine("switching to buffer 1. waited for " + watch.ElapsedMilliseconds);
                currentBufferChars = reloadTask.Result;
                if (currentBufferChars == 0) // reached end of file
                    return '\0';
                if (usingBuffer1)
                {
                    
                    currentBuffer = charBuffer2;
                    standByBuffer = charBuffer1;
                    // reload the finished buffer
                }
                else // using buffer 2
                {
                    
                    currentBuffer = charBuffer1;
                    standByBuffer = charBuffer2;
                    // reload the finished buffer
                }
                usingBuffer1 = !usingBuffer1;
                index = 0;
                //reloadTask = streamReader.ReadAsync(standByBuffer, 0, bufferSize);
                if(streamReader != null)
                    reloadTask = Task.Run<int>(()=> streamReader.Read(standByBuffer, 0, bufferSize));
            }
            return currentBuffer[index++];
        }

        /// <summary>
        /// This method returns a Enumerable of char. Seems to be slower than the GetNextChar()
        /// </summary>
        /// <returns></returns>
        public IEnumerable<char> GetEnumerator()
        {
            while (true)
            {
                if (index >= currentBufferChars) // current buffer is done. switch to next
                {
                    // switch buffer

                    //watch.Restart();
                    reloadTask.Wait();
                    //if (usingBuffer1)
                    //    Console.WriteLine("switching to buffer 2. waited for "+watch.ElapsedMilliseconds);
                    //else
                    //    Console.WriteLine("switching to buffer 1. waited for " + watch.ElapsedMilliseconds);
                    currentBufferChars = reloadTask.Result;
                    if (currentBufferChars == 0) // reached end of file
                        break;
                    if (usingBuffer1)
                    {

                        currentBuffer = charBuffer2;
                        standByBuffer = charBuffer1;
                        // reload the finished buffer
                    }
                    else // using buffer 2
                    {

                        currentBuffer = charBuffer1;
                        standByBuffer = charBuffer2;
                        // reload the finished buffer
                    }
                    usingBuffer1 = !usingBuffer1;
                    index = 0;
                    //reloadTask = streamReader.ReadAsync(standByBuffer, 0, bufferSize);
                    if (streamReader != null)
                        reloadTask = Task.Run<int>(() => streamReader.Read(standByBuffer, 0, bufferSize));
                }
                yield return currentBuffer[index++];
            }
        }

        public BufferedRader(Stream stream, int bufferSize = 10485760)
        {
            this.bufferSize = bufferSize;
            charBuffer1 = new char[this.bufferSize];
            charBuffer2 = new char[this.bufferSize];
            streamReader = new StreamReader(stream, Encoding.UTF8, true, bufferSize);
            currentBufferChars = streamReader.Read(charBuffer1, 0, this.bufferSize);
            currentBuffer = charBuffer1;
            standByBuffer = charBuffer2;
            usingBuffer1 = true;
            reloadTask = Task.Run<int>(() => streamReader.Read(standByBuffer, 0, bufferSize));
        }



        // This part implements the IDisposable interface
        private bool disposed = false;

        public void Dispose()
        {
            reloadTask.Wait();
            reloadTask.Dispose();
            charBuffer1 = new char[0];
            charBuffer2 = charBuffer1;
            currentBuffer = charBuffer1;
            standByBuffer = charBuffer1;
            disposed = true;
            streamReader.Dispose();
        }
    }

    sealed class ExposedBufferedRader : IDisposable
    {
        private StreamReader streamReader;

        private char[] charBuffer1, charBuffer2;

        private char[] currentBuffer, standByBuffer;

        public char[] CurrentBuffer
        {
            get => currentBuffer;
        }

        private int index;

        public int Index
        {
            get => index;

            set
            {
                index = value;
                if(index >= currentBufferChars) // Current buffer ended. Switch buffer
                {
                    reloadTask.Wait();
                    currentBufferChars = reloadTask.Result;
                    if (currentBufferChars == 0)
                    {
                        fileEnded = true;
                        currentBuffer = new char[1] { '\0' };
                        index = 0;
                        return;
                    }
                    if (usingBuffer1)
                    {

                        currentBuffer = charBuffer2;
                        standByBuffer = charBuffer1;
                        // reload the finished buffer
                    }
                    else // using buffer 2
                    {

                        currentBuffer = charBuffer1;
                        standByBuffer = charBuffer2;
                        // reload the finished buffer
                    }
                    usingBuffer1 = !usingBuffer1;
                    index = 0;
                    //reloadTask = streamReader.ReadAsync(standByBuffer, 0, bufferSize);
                    if (streamReader != null)
                        reloadTask = Task.Run<int>(() => streamReader.Read(standByBuffer, 0, bufferSize));
                }
            }
        }

        private bool usingBuffer1;

        public bool fileEnded = false;

        private Task<int> reloadTask;

        public readonly int bufferSize;

        public int currentBufferChars;

        //private System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();

        public char GetNextChar()
        {
            if (index >= currentBufferChars) // current buffer is done. switch to next
            {
                // switch buffer

                //watch.Restart();
                reloadTask.Wait();
                //if (usingBuffer1)
                //    Console.WriteLine("switching to buffer 2. waited for "+watch.ElapsedMilliseconds);
                //else
                //    Console.WriteLine("switching to buffer 1. waited for " + watch.ElapsedMilliseconds);
                currentBufferChars = reloadTask.Result;
                if (currentBufferChars == 0) // reached end of file
                    return '\0';
                if (usingBuffer1)
                {

                    currentBuffer = charBuffer2;
                    standByBuffer = charBuffer1;
                    // reload the finished buffer
                }
                else // using buffer 2
                {

                    currentBuffer = charBuffer1;
                    standByBuffer = charBuffer2;
                    // reload the finished buffer
                }
                usingBuffer1 = !usingBuffer1;
                index = 0;
                //reloadTask = streamReader.ReadAsync(standByBuffer, 0, bufferSize);
                if (streamReader != null)
                    reloadTask = Task.Run<int>(() => streamReader.Read(standByBuffer, 0, bufferSize));
            }
            return currentBuffer[index++];
        }
        

        public ExposedBufferedRader(Stream stream, int bufferSize = 10485760)
        {
            this.bufferSize = bufferSize;
            charBuffer1 = new char[this.bufferSize];
            charBuffer2 = new char[this.bufferSize];
            streamReader = new StreamReader(stream, Encoding.UTF8, true, bufferSize);
            currentBufferChars = streamReader.Read(charBuffer1, 0, this.bufferSize);
            currentBuffer = charBuffer1;
            standByBuffer = charBuffer2;
            usingBuffer1 = true;
            reloadTask = Task.Run<int>(() => streamReader.Read(standByBuffer, 0, bufferSize));
        }



        // This part implements the IDisposable interface
        private bool disposed = false;

        public void Dispose()
        {
            reloadTask.Wait();
            reloadTask.Dispose();
            charBuffer1 = new char[0];
            charBuffer2 = charBuffer1;
            currentBuffer = charBuffer1;
            standByBuffer = charBuffer1;
            disposed = true;
            streamReader.Dispose();
        }
    }

}
