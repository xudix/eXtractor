using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eXtractor
{
    class XlsxReader : IDisposable
    {
        private StreamReader streamReader;

        private char[] charBuffer1, charBuffer2;

        private char[] currentBuffer, standByBuffer;

        private int index = 0;

        private bool usingBuffer1, fileEnded = false;

        private Task<int> reloadTask;

        public readonly int bufferSize;

        public int currentBufferChars;

        /// <summary>
        /// Read a Sheet of a Excel XLSX File. 
        /// </summary>
        /// <param name="stream">A file stream to the sheet.xml file of the XLSX package</param>
        /// <param name="bufferSize">Size of internal char buffers. Default value is 10M chars</param>
        public XlsxReader(Stream stream, int bufferSize = 10485760)
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

        private char[] textBuffer = new char[16384];
        private char[] attributeBuffer = new char[1024];
        private char[] elementNameBuffer = new char[1024];
        private enum ReaderLocationType { StartOfSearch, StartElement, Attribute, AttributeValue, Text, EndElement, EndOfFile }

        /// <summary>
        /// Get the next "row" in the xml file and return it as a string.
        /// </summary>
        /// <returns>A string containing a "row" in the xml file</returns>
        public string GetNextRow()
        {
            char c;
            int textWriteCount = 0, elementNameWriteCount = 0;
            bool isInWantedElement = false;
            ReaderLocationType locationType = ReaderLocationType.StartOfSearch;
            while (!fileEnded)
            {
                c = currentBuffer[index];
                if (isInWantedElement) // Currently in the wanted element. Will write down the current char and watch if we get to the end of the element
                {
                    textBuffer[textWriteCount] = c;
                    textWriteCount++;
                    if (textWriteCount == textBuffer.Length)
                        Array.Resize(ref textBuffer, textBuffer.Length + 4096);
                    switch (locationType)
                    {
                        case ReaderLocationType.Text:
                            if (c == '<')
                            {
                                locationType = ReaderLocationType.StartElement;
                            }
                            break;
                        case ReaderLocationType.StartElement:
                            if (c == '/') // it's an EndElement
                                locationType = ReaderLocationType.EndElement;
                            else
                                locationType = ReaderLocationType.Text;
                            break;
                        case ReaderLocationType.EndElement: // need to determine if this is the end of "elementWanted"
                            switch (c)
                            {
                                case ' ': // end of element name. should not see this in endElement
                                    if (elementNameWriteCount == 3 && elementNameBuffer[0] == 'r'
                                        && elementNameBuffer[1] == 'o' && elementNameBuffer[2] == 'w')// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</ " in the EndElement
                                            return new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        else
                                            return String.Empty;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.Text;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '>':
                                    if (elementNameWriteCount == 3 && elementNameBuffer[0] == 'r'
                                        && elementNameBuffer[1] == 'o' && elementNameBuffer[2] == 'w')// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</ " in the EndElement
                                            return new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        else
                                            return String.Empty;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.Text;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                default:
                                    elementNameBuffer[elementNameWriteCount] = c;
                                    elementNameWriteCount++;
                                    break;
                            }
                            break;
                    }
                }
                else // Not in a row. Will not write down the current char. Watch for start element of row
                {
                    switch (locationType) // The following code does not write anything to textBuffer. It should only write to attributes and control the locationType
                    {
                        case ReaderLocationType.StartOfSearch:
                            if (c == '<')
                            {
                                locationType = ReaderLocationType.StartElement;
                            }
                            break;
                        case ReaderLocationType.StartElement:
                            switch (c)
                            {
                                case ' ': // end of element name. Upcoming things will be attributes
                                    if (elementNameWriteCount == 3 && elementNameBuffer[0] == 'r'
                                        && elementNameBuffer[1] == 'o' && elementNameBuffer[2] == 'w')// got into the element we want. 
                                    {
                                        // Don't care about the attribute here. But this signals that we got into a row
                                        locationType = ReaderLocationType.Attribute;
                                    }
                                    else // other element and not inside a row
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    if (elementNameWriteCount == 3 && elementNameBuffer[0] == 'r'
                                        && elementNameBuffer[1] == 'o' && elementNameBuffer[2] == 'w')// got into the element we want.
                                    {
                                        locationType = ReaderLocationType.Text;
                                        isInWantedElement = true;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        if (elementNameWriteCount == 3 && elementNameBuffer[0] == 'r'
                                        && elementNameBuffer[1] == 'o' && elementNameBuffer[2] == 'w')// got into the element we want. Will return
                                        {
                                            if (textWriteCount > 0)
                                                return new string(textBuffer, 0, textWriteCount);
                                            else
                                                return String.Empty;
                                        }
                                        else // other element and not inside the wanted element.
                                        {
                                            // search for the next element
                                            locationType = ReaderLocationType.StartOfSearch;
                                        }
                                        elementNameWriteCount = 0;
                                    }
                                    else // it's following a < sign. End element. Since we are not in the wanted element, don't care about this element.
                                        locationType = ReaderLocationType.StartOfSearch;
                                    break;
                                default:
                                    elementNameBuffer[elementNameWriteCount] = c;
                                    elementNameWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.Attribute: // Attribute only exis in the StartElement. 
                            switch (c)
                            {
                                case '/': // end of attribute and the whole element. should return.
                                    if (textWriteCount > 0)
                                        return new string(textBuffer, 0, textWriteCount);
                                    else
                                        return String.Empty;
                                case '>': // end of attribute and start element. start text part.
                                    locationType = ReaderLocationType.Text;
                                    isInWantedElement = true;
                                    break;
                                default:
                                    break;
                            }
                            break;
                        case ReaderLocationType.Text: // Shouldn't get here
                            if (c == '<')
                            {
                                locationType = ReaderLocationType.StartElement;
                            }
                            break;
                    }
                }

                // move to the next char.
                index++;
                if (index >= currentBufferChars) // Current buffer ended. Switch buffer
                {
                    reloadTask.Wait();
                    currentBufferChars = reloadTask.Result;
                    if (currentBufferChars == 0)
                    {
                        fileEnded = true;
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
            // get to the end of the file
            if (textWriteCount > 0)
                return new string(textBuffer, 0, textWriteCount);
            else
                return String.Empty;
        }

        public void Dispose()
        {
            reloadTask.Wait();
            reloadTask.Dispose();
            charBuffer1 = new char[0];
            charBuffer2 = charBuffer1;
            currentBuffer = charBuffer1;
            standByBuffer = charBuffer1;
            streamReader.Dispose();
        }
    }
}
