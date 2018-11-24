using System;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace eXtractor
{
    /// <summary>
    /// Tools to process Excel OpenXML (XLSX) files
    /// </summary>
    class XlsxTool
    {
        /// <summary>
        /// Get the shared string table of the Excel OpenXML (XLSX) file and return it in an string array
        /// </summary>
        /// <param name="fileName">The file name of the Excel OpenXML (XLSX) file containing the header</param>
        /// <returns>The header (first row) of the Excel OpenXML (XLSX) file</returns>
        public static string[] GetSharedStrings(string fileName)
            => GetSharedStrings(new ZipArchive(new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)));

        /// <summary>
        /// Get the shared string table of the Excel OpenXML (XLSX) file and return it in an string array
        /// </summary>
        /// <param name="zipArchive">The file stream of the Excel OpenXML (XLSX) file containing the header</param>
        /// <returns>The header (first row) of the Excel OpenXML (XLSX) file</returns>
        public static string[] GetSharedStrings(ZipArchive zipArchive)
        {
            SharedStringTable SharedStrings;
            string[] result;
            SharedStrings = DeserializedZipEntry<SharedStringTable>(zipArchive.GetEntry(@"xl/sharedStrings.xml"));
            result = new string[Int32.Parse(SharedStrings.count)];
            for (int i = 0; i < SharedStrings.si.Length; i++)
                result[i] = SharedStrings.si[i].t;
            return result;
        }

        /// <summary>
        /// Get the header (first row) of the first sheet of an Excel OpenXML (XLSX) file
        /// </summary>
        /// <param name="zipArchive"></param>
        /// <returns></returns>
        public static string[] GetHeader(ZipArchive zipArchive)
        {
            // Normally the sharedStrings should be all the headers we need. 
            // However, in principal, if the file contains other strings, the shared string table is larger than the first row
            // Thus, we open the sheet to actually process the first row.
            string[] sharedStrings = GetSharedStrings(zipArchive);
            
            XmlEntry firstRow;
            int sharedStringIndex;
            // open the first sheet of the file
            using (StreamReader worksheetReader = new StreamReader(zipArchive.GetEntry(@"xl/worksheets/sheet1.xml").Open()))
            {
                firstRow = XlsxReadOne(worksheetReader, "row");
            }
            // method one: RegEx
            MatchCollection matches = Regex.Matches(firstRow.text, @"(?<=<v>)(.*?)(?=<)");
            string[] result = new string[matches.Count];
            for (int i = 0; i < matches.Count; i++)
            {
                sharedStringIndex = Int32.Parse(matches[i].Value);
                result[i] = sharedStrings[sharedStringIndex];
            }
            // method two: XlsxReadOne(stringreader, "v")
            // XmlEntry cell, value;
            //using(StringReader sr = new StringReader(firstRow.text))
            //{
            //    while (sr.Peek() >= 0)
            //    {
            //        cell = XlsxReadOne(sr, "c");

            //    }
            //}

            return result;
        }

        /// <summary>
        /// Get the header (first row) and the corresponding column references of the first sheet of an Excel OpenXML (XLSX) file
        /// </summary>
        /// <param name="zipArchive"></param>
        /// <returns></returns>
        public static HeaderWithColRef GetHeaderWithColReference(ZipArchive zipArchive)
        {
            // Normally the sharedStrings should be all the headers we need. 
            // However, in principal, if the file contains other strings, the shared string table is larger than the first row
            // Thus, we open the sheet to actually process the first row.
            string[] sharedStrings = GetSharedStrings(zipArchive);
            HeaderWithColRef result;
            XmlEntry firstRow;
            int sharedStringIndex;
            // open the first sheet of the file
            using (StreamReader worksheetReader = new StreamReader(zipArchive.GetEntry(@"xl/worksheets/sheet1.xml").Open()))
            {
                firstRow = XlsxReadOne(worksheetReader, "row");
            }
            // method one: RegEx
            string pattern = @"<c\s*(?:r=""(?<colRef>[A-Z]+)[0-9]+"")?(?:\s(?:cm|ph|s|vm)="".*?"")*(?:\st=""(?<type>[a-z]*)"")?.*?>\s*<v>(?<value>.*?)(?=<)";
            MatchCollection matches = Regex.Matches(firstRow.text, pattern);
            result.header = new string[matches.Count];
            result.colRef = new string[matches.Count];
            for (int i = 0; i < matches.Count; i++)
            {
                if (matches[i].Groups["type"].Value == "s")// The cell contains string data, which should come from shared string table
                {
                    sharedStringIndex = Int32.Parse(matches[i].Groups["value"].Value);
                    result.header[i] = sharedStrings[sharedStringIndex];
                }
                else
                    result.header[i] = matches[i].Groups["value"].Value;
                result.colRef[i] = matches[i].Groups["colRef"].Value;
            }
            // method two: XlsxReadOne(stringreader, "v")
            // XmlEntry cell, value;
            //using(StringReader sr = new StringReader(firstRow.text))
            //{
            //    while (sr.Peek() >= 0)
            //    {
            //        cell = XlsxReadOne(sr, "c");

            //    }
            //}

            return result;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="elementWanted"></param>
        /// <returns></returns>
        public static XmlEntry XlsxReadOne(TextReader reader, string elementWanted)
        {
            string elementName;
            char c;
            int c_int;
            char[] textBuffer = new char[16384];
            char[] attributeBuffer = new char[1024];
            char[] elementNameBuffer = new char[1024];
            int textWriteCount = 0, attrWriteCount = 0, elementNameWriteCount = 0;
            bool isInWantedElement = false;
            ReaderLocationType locationType = ReaderLocationType.StartOfSearch;
            XmlEntry result;
            result.text = String.Empty;
            result.xmlAttributes = new List<XmlAttributeItem>();
            while((c_int = reader.Read())>=0)
            {
                c = (char)c_int;
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
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    if (elementName == elementWanted)// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</ " in the EndElement
                                            result.text = new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        return result;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.Text;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '>':
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    if (elementName == elementWanted)// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</>" in the EndElement
                                            result.text = new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        return result;
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
                else // Not in the wanted element. Will not write down the current char. Watch for start element of wanted char
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
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == elementWanted)// got into the element we want. Will read the attribute 
                                    {
                                        locationType = ReaderLocationType.Attribute;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == elementWanted)// got into the element we want.
                                    {
                                        locationType = ReaderLocationType.Text;
                                        isInWantedElement = true;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                        elementNameWriteCount = 0;
                                        if (elementName == elementWanted)// got into the element we want. Will return
                                        {
                                            if (textWriteCount > 0)
                                                result.text = new string(textBuffer, 0, textWriteCount);
                                            return result;
                                        }
                                        else // other element and not inside the wanted element.
                                        {
                                            // search for the next element
                                            locationType = ReaderLocationType.StartOfSearch;
                                        }
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
                                case ' ': // end of an attribute. Another attribute coming.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    attrWriteCount = 0;
                                    break;
                                case '/': // end of attribute and the whole element. should return.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    if (textWriteCount > 0)
                                        result.text = new string(textBuffer, 0, textWriteCount);
                                    attrWriteCount = 0;
                                    return result;
                                case '>': // end of attribute and start element. start text part.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    attrWriteCount = 0;
                                    locationType = ReaderLocationType.Text;
                                    isInWantedElement = true;
                                    break;
                                default:
                                    attributeBuffer[attrWriteCount] = c;
                                    attrWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.Text: // seems to be the same as "StartOfSearch"
                            if (c == '<')
                            {
                                locationType = ReaderLocationType.StartElement;
                            }
                            break;
                    }
                }
            }
            // get to the end of the file
            if (textWriteCount > 0)
                result.text = new string(textBuffer, 0, textWriteCount);
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="elementWanted"></param>
        /// <returns></returns>
        public static XmlEntry XlsxReadOneFromBufferedReader(BufferedRader reader, string elementWanted)
        {
            string elementName;
            char c;
            char[] textBuffer = new char[16384];
            char[] attributeBuffer = new char[1024];
            char[] elementNameBuffer = new char[1024];
            int textWriteCount = 0, attrWriteCount = 0, elementNameWriteCount = 0;
            bool isInWantedElement = false;
            ReaderLocationType locationType = ReaderLocationType.StartOfSearch;
            XmlEntry result;
            result.text = String.Empty;
            result.xmlAttributes = new List<XmlAttributeItem>();
            while ((c = reader.GetNextChar()) > 0)
            //foreach(char c in reader.GetEnumerator())
            {
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
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    if (elementName == elementWanted)// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</ " in the EndElement
                                            result.text = new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        return result;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.Text;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '>':
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    if (elementName == elementWanted)// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</>" in the EndElement
                                            result.text = new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        return result;
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
                else // Not in the wanted element. Will not write down the current char. Watch for start element of wanted char
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
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == elementWanted)// got into the element we want. Will read the attribute 
                                    {
                                        locationType = ReaderLocationType.Attribute;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == elementWanted)// got into the element we want.
                                    {
                                        locationType = ReaderLocationType.Text;
                                        isInWantedElement = true;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                        elementNameWriteCount = 0;
                                        if (elementName == elementWanted)// got into the element we want. Will return
                                        {
                                            if (textWriteCount > 0)
                                                result.text = new string(textBuffer, 0, textWriteCount);
                                            return result;
                                        }
                                        else // other element and not inside the wanted element.
                                        {
                                            // search for the next element
                                            locationType = ReaderLocationType.StartOfSearch;
                                        }
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
                                case ' ': // end of an attribute. Another attribute coming.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    attrWriteCount = 0;
                                    break;
                                case '/': // end of attribute and the whole element. should return.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    if (textWriteCount > 0)
                                        result.text = new string(textBuffer, 0, textWriteCount);
                                    attrWriteCount = 0;
                                    return result;
                                case '>': // end of attribute and start element. start text part.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    attrWriteCount = 0;
                                    locationType = ReaderLocationType.Text;
                                    isInWantedElement = true;
                                    break;
                                default:
                                    attributeBuffer[attrWriteCount] = c;
                                    attrWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.Text: // seems to be the same as "StartOfSearch"
                            if (c == '<')
                            {
                                locationType = ReaderLocationType.StartElement;
                            }
                            break;
                    }
                }
            }
            // get to the end of the file
            if (textWriteCount > 0)
                result.text = new string(textBuffer, 0, textWriteCount);
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="elementWanted"></param>
        /// <returns></returns>
        public static XmlEntry XlsxReadOneFromExposedBufferedReader(ExposedBufferedRader reader, string elementWanted)
        {
            string elementName;
            char c;
            char[] textBuffer = new char[16384];
            char[] attributeBuffer = new char[1024];
            char[] elementNameBuffer = new char[1024];
            int textWriteCount = 0, attrWriteCount = 0, elementNameWriteCount = 0;
            bool isInWantedElement = false;
            ReaderLocationType locationType = ReaderLocationType.StartOfSearch;
            XmlEntry result;
            result.text = String.Empty;
            result.xmlAttributes = new List<XmlAttributeItem>();
            while (!reader.fileEnded)
            //foreach(char c in reader.GetEnumerator())
            {
                c = reader.CurrentBuffer[reader.Index++];
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
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    if (elementName == elementWanted)// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</ " in the EndElement
                                            result.text = new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        return result;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.Text;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '>':
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    if (elementName == elementWanted)// finished the element we want. Will return
                                    {
                                        if (textWriteCount > 0)
                                            // The number of chars written to the string is reduced to takeout:
                                            // the element name and "</>" in the EndElement
                                            result.text = new string(textBuffer, 0, textWriteCount - elementNameWriteCount - 3);
                                        return result;
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
                else // Not in the wanted element. Will not write down the current char. Watch for start element of wanted char
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
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == elementWanted)// got into the element we want. Will read the attribute 
                                    {
                                        locationType = ReaderLocationType.Attribute;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == elementWanted)// got into the element we want.
                                    {
                                        locationType = ReaderLocationType.Text;
                                        isInWantedElement = true;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                        elementNameWriteCount = 0;
                                        if (elementName == elementWanted)// got into the element we want. Will return
                                        {
                                            if (textWriteCount > 0)
                                                result.text = new string(textBuffer, 0, textWriteCount);
                                            return result;
                                        }
                                        else // other element and not inside the wanted element.
                                        {
                                            // search for the next element
                                            locationType = ReaderLocationType.StartOfSearch;
                                        }
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
                                case ' ': // end of an attribute. Another attribute coming.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    attrWriteCount = 0;
                                    break;
                                case '/': // end of attribute and the whole element. should return.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    if (textWriteCount > 0)
                                        result.text = new string(textBuffer, 0, textWriteCount);
                                    attrWriteCount = 0;
                                    return result;
                                case '>': // end of attribute and start element. start text part.
                                    result.xmlAttributes.Add(new XmlAttributeItem(new string(attributeBuffer, 0, attrWriteCount)));
                                    attrWriteCount = 0;
                                    locationType = ReaderLocationType.Text;
                                    isInWantedElement = true;
                                    break;
                                default:
                                    attributeBuffer[attrWriteCount] = c;
                                    attrWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.Text: // seems to be the same as "StartOfSearch"
                            if (c == '<')
                            {
                                locationType = ReaderLocationType.StartElement;
                            }
                            break;
                    }
                }
            }
            // get to the end of the file
            if (textWriteCount > 0)
                result.text = new string(textBuffer, 0, textWriteCount);
            return result;
        }


        public static T GetValueFromRow<T>(string row, string colRef)
        {
            string elementName, valueText = "";
            char c;
            char[] textBuffer = new char[4096];
            char[] attributeBuffer = new char[128];
            char[] elementNameBuffer = new char[128];
            int textWriteCount = 0, attrWriteCount = 0, elementNameWriteCount = 0;
            bool isInWantedCell = false;
            ReaderLocationType locationType = ReaderLocationType.StartOfSearch;
            for(int i = 0; i < row.Length; i++)
            {
                c = row[i];
                if (isInWantedCell) // Currently cursor in the wanted cell
                {
                    switch (locationType)
                    {
                        case ReaderLocationType.StartOfSearch:
                            if (c == '<')
                                locationType = ReaderLocationType.StartElement;
                            break;
                        case ReaderLocationType.StartElement:
                            switch (c)
                            {
                                case ' ': // end of element name, followed by attribute. should not get here
                                    if (elementNameBuffer[0] != 'v' || elementNameWriteCount != 1)// Not the value part. will skip this element
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    else // if for some reason the value part have attribute. Don't need to know the attribute for now
                                        locationType = ReaderLocationType.Attribute;
                                    elementNameWriteCount = 0;
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    if (elementNameBuffer[0] == 'v' || elementNameWriteCount == 1)// got into the value part.
                                    {
                                        locationType = ReaderLocationType.Text;
                                    }
                                    else // other element
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        if (elementNameBuffer[0] == 'v' || elementNameWriteCount == 1)// the value part is empty.
                                        {
                                            // prepare to return. Set cursor to the end of the row
                                            i = row.Length;
                                        }
                                        else // other element
                                        {
                                            // search for the next element
                                            locationType = ReaderLocationType.StartOfSearch;
                                        }
                                    }
                                    else // it's following a < sign. End element. Not the v element. Keep searching
                                        locationType = ReaderLocationType.StartOfSearch;
                                    break;
                                default:
                                    elementNameBuffer[elementNameWriteCount] = c;
                                    elementNameWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.Attribute:
                            if (c == '>')//end of attribute of 'v' part. will start recording the text
                            {
                                locationType = ReaderLocationType.Text;
                            }
                            break;
                        case ReaderLocationType.Text:
                            switch (c)
                            {
                                case '<': // finished reading the 'v' part. will return. set the cursor to the end of the row string
                                    i = row.Length;
                                    break;
                                default:
                                    textBuffer[textWriteCount] = c;
                                    textWriteCount++;
                                    break;
                            }
                            break;
                        default:
                            throw new InvalidDataException("Fail to process row.\n" + row);
                    }
                }
                else // currently cursor not in the wanted cell
                {
                    switch (locationType)
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
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == "c")// got into a cell. Will read the attribute 
                                    {
                                        locationType = ReaderLocationType.Attribute;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                    elementNameWriteCount = 0;
                                    if (elementName == "c")// got into a cell. This should not happen since there should be attribute
                                    {
                                        throw (new InvalidDataException("The cell does not contain Reference Attribute.\n" + row));
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        elementName = new string(elementNameBuffer, 0, elementNameWriteCount);
                                        elementNameWriteCount = 0;
                                        if (elementName == "c")// got into a cell. This should not happen since there should be attribute
                                        {
                                            throw (new InvalidDataException("The cell does not contain Reference Attribute.\n" + row));
                                        }
                                        else // other element and not inside the wanted element.
                                        {
                                            // search for the next element
                                            locationType = ReaderLocationType.StartOfSearch;
                                        }
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
                        case ReaderLocationType.Attribute:// Attribute only exis in the StartElement. 
                            switch (c)
                            {
                                case '=': // end of attribute name
                                    if (attributeBuffer[0] == 'r' && attrWriteCount == 1)// found the Reference attribute
                                    {
                                        // Will read the attribute value
                                        locationType = ReaderLocationType.AttributeValue;
                                    }
                                    attrWriteCount = 0;
                                    break;
                                case ' ': // end of an attribute. Another attribute coming.
                                    attrWriteCount = 0;
                                    break;
                                case '/': // end of attribute and the whole element. Empty cell?
                                    throw new InvalidDataException("The cell does not contain Reference Attribute.\n" + row);
                                case '>': // end of attribute and start element. Should not get here
                                    throw new InvalidDataException("The cell does not contain Reference Attribute.\n" + row);
                                default:
                                    attributeBuffer[attrWriteCount] = c;
                                    attrWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.AttributeValue:
                            switch (c)
                            {
                                case '\"':
                                    if(attrWriteCount > 0)//finished reading the reference attribute. Should not get here, since only column reference is needed
                                    {
                                        throw new InvalidDataException("The cell contains invalid Reference Attribute.\n" + row);
                                    }
                                    break;
                                case ' ':
                                case '/':
                                case '>': //should not see these.
                                    throw new InvalidDataException("The cell contains invalid Reference Attribute.\n" + row);
                                default:
                                    if (c >= '0' && c <= '9') // end of column reference
                                    {
                                        if(new string(attributeBuffer, 0, attrWriteCount) == colRef) // this is the column we need
                                        {
                                            isInWantedCell = true;
                                            
                                        }// If is not the column we need, switch back to StartOfSearch
                                        attrWriteCount = 0;
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    else // still reading column reference
                                    {
                                        attributeBuffer[attrWriteCount] = c;
                                        attrWriteCount++;
                                    }
                                    break;
                            }
                            break;
                        default:
                            throw new InvalidDataException("Fail to process row.\n" + row);
                    }
                }
            }
            if (textWriteCount > 0)
                valueText = new string(textBuffer, 0, textWriteCount);
            return (T)Convert.ChangeType(valueText, typeof(T));

        }


        public static double GetDoubleFromRow(string row, string colRef)
        {
            char c;
            char[] textBuffer = new char[4096];
            char[] attributeBuffer = new char[128];
            char[] elementNameBuffer = new char[128];
            int textWriteCount = 0, attrWriteCount = 0, elementNameWriteCount = 0;
            bool isInWantedCell = false;
            ReaderLocationType locationType = ReaderLocationType.StartOfSearch;
            for (int i = 0; i < row.Length; i++)
            {
                c = row[i];
                if (isInWantedCell) // Currently cursor in the wanted cell
                {
                    switch (locationType)
                    {
                        case ReaderLocationType.StartOfSearch:
                            if (c == '<')
                                locationType = ReaderLocationType.StartElement;
                            break;
                        case ReaderLocationType.StartElement:
                            switch (c)
                            {
                                case ' ': // end of element name, followed by attribute. should not get here
                                    if (elementNameBuffer[0] != 'v' || elementNameWriteCount != 1)// Not the value part. will skip this element
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    else // if for some reason the value part have attribute. Don't need to know the attribute for now
                                        locationType = ReaderLocationType.Attribute;
                                    elementNameWriteCount = 0;
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    if (elementNameBuffer[0] == 'v' || elementNameWriteCount == 1)// got into the value part.
                                    {
                                        locationType = ReaderLocationType.Text;
                                    }
                                    else // other element
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        if (elementNameBuffer[0] == 'v' || elementNameWriteCount == 1)// the value part is empty.
                                        {
                                            // prepare to return. Set cursor to the end of the row
                                            i = row.Length;
                                        }
                                        else // other element
                                        {
                                            // search for the next element
                                            locationType = ReaderLocationType.StartOfSearch;
                                        }
                                    }
                                    else // it's following a < sign. End element. Not the v element. Keep searching
                                        locationType = ReaderLocationType.StartOfSearch;
                                    break;
                                default:
                                    elementNameBuffer[elementNameWriteCount] = c;
                                    elementNameWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.Attribute:
                            if (c == '>')//end of attribute of 'v' part. will start recording the text
                            {
                                locationType = ReaderLocationType.Text;
                            }
                            break;
                        case ReaderLocationType.Text:
                            switch (c)
                            {
                                case '<': // finished reading the 'v' part. will return. set the cursor to the end of the row string
                                    i = row.Length;
                                    break;
                                default:
                                    textBuffer[textWriteCount] = c;
                                    textWriteCount++;
                                    break;
                            }
                            break;
                        default:
                            throw new InvalidDataException("Fail to process row.\n" + row);
                    }
                }
                else // currently cursor not in the wanted cell
                {
                    switch (locationType)
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
                                    if (elementNameBuffer[0] == 'c' && elementNameWriteCount == 1)// got into a cell. Will read the attribute 
                                    {
                                        locationType = ReaderLocationType.Attribute;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    if (elementNameBuffer[0] == 'c' && elementNameWriteCount == 1)// got into a cell. This should not happen since there should be attribute
                                    {
                                        throw (new InvalidDataException("The cell does not contain Reference Attribute.\n" + row));
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
                                        if (elementNameBuffer[0] == 'c' && elementNameWriteCount == 1)// got into a cell. This should not happen since there should be attribute
                                        {
                                            throw (new InvalidDataException("The cell does not contain Reference Attribute.\n" + row));
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
                        case ReaderLocationType.Attribute:// Attribute only exis in the StartElement. 
                            switch (c)
                            {
                                case '=': // end of attribute name
                                    if (attributeBuffer[0] == 'r' && attrWriteCount == 1)// found the Reference attribute
                                    {
                                        // Will read the attribute value
                                        locationType = ReaderLocationType.AttributeValue;
                                    }
                                    attrWriteCount = 0;
                                    break;
                                case ' ': // end of an attribute. Another attribute coming.
                                    attrWriteCount = 0;
                                    break;
                                case '/': // end of attribute and the whole element. Empty cell?
                                    throw new InvalidDataException("The cell does not contain Reference Attribute.\n" + row);
                                case '>': // end of attribute and start element. Should not get here
                                    throw new InvalidDataException("The cell does not contain Reference Attribute.\n" + row);
                                default:
                                    attributeBuffer[attrWriteCount] = c;
                                    attrWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.AttributeValue:
                            switch (c)
                            {
                                case '\"':
                                    if (attrWriteCount > 0)//finished reading the reference attribute. Should not get here, since only column reference is needed
                                    {
                                        throw new InvalidDataException("The cell contains invalid Reference Attribute.\n" + row);
                                    }
                                    break;
                                case ' ':
                                case '/':
                                case '>': //should not see these.
                                    throw new InvalidDataException("The cell contains invalid Reference Attribute.\n" + row);
                                default:
                                    if (c >= '0' && c <= '9') // end of column reference
                                    {
                                        if (new string(attributeBuffer, 0, attrWriteCount) == colRef) // this is the column we need
                                        {
                                            isInWantedCell = true;

                                        }// If is not the column we need, switch back to StartOfSearch
                                        attrWriteCount = 0;
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    else // still reading column reference
                                    {
                                        attributeBuffer[attrWriteCount] = c;
                                        attrWriteCount++;
                                    }
                                    break;
                            }
                            break;
                        default:
                            throw new InvalidDataException("Fail to process row.\n" + row);
                    }
                }
            }
            if (textWriteCount > 0)
                return Double.Parse(new string(textBuffer, 0, textWriteCount));
            else
                return Double.NaN;

        }


        public static void GetFloatsFromRow(string row, IList<string> colRef, ref float[] result)
        {
            char c;
            char[] textBuffer = new char[4096];
            char[] attributeBuffer = new char[128];
            char[] elementNameBuffer = new char[128];
            int textWriteCount = 0, attrWriteCount = 0, elementNameWriteCount = 0;
            int resultCount = 0; // number of items found in colRef
            int colNumber; // The colume number corresponding to the A1 style column index. obtained from the ConvertColRef method.
            bool isInWantedCell = false;
            ReaderLocationType locationType = ReaderLocationType.StartOfSearch;
            if (String.IsNullOrEmpty(colRef[0]))
            {
                for (int i = resultCount; i < colRef.Count; i++)
                    result[i] = Single.NaN;
                return;
            }
            colNumber = ConvertColRef(colRef[0]);
            for (int i = 0; i < row.Length; i++)
            {
                c = row[i];
                if (isInWantedCell) // Currently cursor in the wanted cell. will search for 'v' element and read the text in it.
                {
                    switch (locationType)
                    {
                        case ReaderLocationType.StartOfSearch:
                            if (c == '<')
                                locationType = ReaderLocationType.StartElement;
                            break;
                        case ReaderLocationType.StartElement:
                            switch (c)
                            {
                                case ' ': // end of element name, followed by attribute. should not get here
                                    if (elementNameBuffer[0] != 'v' || elementNameWriteCount != 1)// Not the value part. will skip this element
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    else // if for some reason the value part have attribute. Don't need to know the attribute for now
                                        locationType = ReaderLocationType.Attribute;
                                    elementNameWriteCount = 0;
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    if (elementNameBuffer[0] == 'v' || elementNameWriteCount == 1)// got into the value part.
                                    {
                                        locationType = ReaderLocationType.Text;
                                    }
                                    else // other element
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '/': // end of element. 
                                    if (elementNameWriteCount != 0) // It's not following the < sign. It's an empty element.
                                    {
                                        if (elementNameBuffer[0] == 'v' || elementNameWriteCount == 1)// the value part is empty.
                                        {
                                            // write NaN to the result array, and continue to the next one
                                            result[resultCount] = Single.NaN;
                                            resultCount++;
                                            if (resultCount == colRef.Count) // found all items requested by colRef
                                                return;
                                            if (String.IsNullOrEmpty(colRef[resultCount]))
                                            {
                                                for (int j = resultCount; j < colRef.Count; j++)
                                                    result[j] = Single.NaN;
                                                return;
                                            }
                                            colNumber = ConvertColRef(colRef[resultCount]);
                                            isInWantedCell = false;
                                        }
                                    }
                                    locationType = ReaderLocationType.StartOfSearch;
                                    break;
                                default:
                                    elementNameBuffer[elementNameWriteCount] = c;
                                    elementNameWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.Attribute:
                            if (c == '>')//end of attribute of 'v' part. will start recording the text
                            {
                                locationType = ReaderLocationType.Text;
                            }
                            break;
                        case ReaderLocationType.Text:
                            switch (c)
                            {
                                case '<': // finished reading the 'v' part. Record the value to result and move to the next one
                                    result[resultCount] = Single.Parse(new string(textBuffer, 0, textWriteCount));
                                    textWriteCount = 0;
                                    resultCount++;
                                    if (resultCount == colRef.Count) // found all items requested by colRef
                                        return;
                                    if (String.IsNullOrEmpty(colRef[resultCount]))
                                    {
                                        for (int j = resultCount; j < colRef.Count; j++)
                                            result[j] = Single.NaN;
                                        return;
                                    }
                                    colNumber = ConvertColRef(colRef[resultCount]);
                                    isInWantedCell = false;
                                    locationType = ReaderLocationType.StartOfSearch;
                                    break;
                                default:
                                    textBuffer[textWriteCount] = c;
                                    textWriteCount++;
                                    break;
                            }
                            break;
                        default:
                            throw new InvalidDataException("Fail to process row.\n" + row);
                    }
                }
                else // currently cursor not in the wanted cell. Will search for 'c' and read 'r' attribute
                {
                    switch (locationType)
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
                                    if (elementNameBuffer[0] == 'c' && elementNameWriteCount == 1)// got into a cell. Will read the attribute 
                                    {
                                        locationType = ReaderLocationType.Attribute;
                                    }
                                    else // other element and not inside the wanted element.
                                    {
                                        // search for the next element
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    elementNameWriteCount = 0;
                                    break;
                                case '>': // end of element name and no attributel go to text.
                                    if (elementNameBuffer[0] == 'c' && elementNameWriteCount == 1)// got into a cell. This should not happen since there should be attribute
                                    {
                                        throw (new InvalidDataException("The cell does not contain Reference Attribute.\n" + row));
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
                                        if (elementNameBuffer[0] == 'c' && elementNameWriteCount == 1)// got into a cell. This should not happen since there should be attribute
                                        {
                                            throw (new InvalidDataException("The cell does not contain Reference Attribute.\n" + row));
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
                        case ReaderLocationType.Attribute:// Attribute only exis in the StartElement. 
                            switch (c)
                            {
                                case '=': // end of attribute name
                                    if (attributeBuffer[0] == 'r' && attrWriteCount == 1)// found the Reference attribute
                                    {
                                        // Will read the attribute value
                                        locationType = ReaderLocationType.AttributeValue;
                                    }
                                    attrWriteCount = 0;
                                    break;
                                case ' ': // end of an attribute. Another attribute coming.
                                    attrWriteCount = 0;
                                    break;
                                case '/': // end of attribute and the whole element. Empty cell?
                                    throw new InvalidDataException("The cell does not contain Reference Attribute.\n" + row);
                                case '>': // end of attribute and start element. Should not get here
                                    throw new InvalidDataException("The cell does not contain Reference Attribute.\n" + row);
                                default:
                                    attributeBuffer[attrWriteCount] = c;
                                    attrWriteCount++;
                                    break;
                            }
                            break;
                        case ReaderLocationType.AttributeValue:
                            switch (c)
                            {
                                case '\"':
                                    if (attrWriteCount > 0)//finished reading the reference attribute. Should not get here, there should be a number in the reference part, which cause isInWantedCell = true.
                                    {
                                        throw new InvalidDataException("The cell contains invalid Reference Attribute.\n" + row);
                                    }
                                    break;
                                case ' ':
                                case '/':
                                case '>': //should not see these.
                                    throw new InvalidDataException("The cell contains invalid Reference Attribute.\n" + row);
                                default:
                                    if (c<='9') // end of column reference
                                    {
                                        int currentColNum = ConvertColRef(new string(attributeBuffer, 0, attrWriteCount));
                                        while(currentColNum > colNumber) // the requested colRef or colNum is not found. Later columns has showed up.
                                        {
                                            // write NaN to the result array, and continue to the next one
                                            result[resultCount] = Single.NaN;
                                            resultCount++;
                                            if (resultCount == colRef.Count) // found all items requested by colRef
                                                return;
                                            if (String.IsNullOrEmpty(colRef[resultCount]))
                                            {
                                                for (int j = resultCount; j < colRef.Count; j++)
                                                    result[j] = Single.NaN;
                                                return;
                                            }
                                            colNumber = ConvertColRef(colRef[resultCount]); // this is the next wanted colRef
                                            isInWantedCell = false;
                                        }
                                        if (currentColNum == colNumber) // this is the column we need
                                        {
                                            isInWantedCell = true;
                                        }
                                        // If is not the column we need, switch back to StartOfSearch
                                        attrWriteCount = 0;
                                        locationType = ReaderLocationType.StartOfSearch;
                                    }
                                    else // still reading column reference
                                    {
                                        attributeBuffer[attrWriteCount] = c;
                                        attrWriteCount++;
                                    }
                                    break;
                            }
                            break;
                        default:
                            throw new InvalidDataException("Fail to process row.\n" + row);
                    }
                }
            }
            // If we didn't return before the for loop ends, it means we finished reading all chars in the row and didn't find all items
            // will write NaN to the rest
            for (int i = resultCount; i < colRef.Count; i++)
                result[i] = Single.NaN;
        }

        /// <summary>
        /// Convert a A1 style column reference (A, B, ... Z, AA, AB, ..., ZZ,...) into a integer corresponding to the column number
        /// "A" => 1, "B" => 2, "Z" => 26, "AA" => 27
        /// </summary>
        /// <param name="colRef">A column reference in A1 style. No number is allowed.</param>
        /// <returns>The column number corresponding to the column reference.</returns>
        public static int ConvertColRef(string colRef)
        {
            int result = 0, i;
            for ( i = 0; i < colRef.Length; i++)
                result += (colRef[i] - 64) * (int)Math.Pow(26, colRef.Length-i-1);
            return result;
        }

        private enum ReaderLocationType {StartOfSearch, StartElement, Attribute, AttributeValue, Text, EndElement, EndOfFile }


        private static T DeserializedZipEntry<T>(ZipArchiveEntry ZipArchiveEntry)
        {
            using (Stream stream = ZipArchiveEntry.Open())
                return (T)new XmlSerializer(typeof(T)).Deserialize(XmlReader.Create(stream));
        }

        public struct XmlEntry
        {
            public string text;
            public List<XmlAttributeItem> xmlAttributes;

            //public string this[string index]
            //{
            //    get
            //    {
                    
            //    }
            //}

        }


        public struct XmlAttributeItem
        {
            string Name;
            string Value;
            
            /// <summary>
            /// Convert the Xml attribute string in to a XmlAttributeItem.
            /// </summary>
            /// <param name="xmlAttributeText">The string containing the attribute in Xml file. It should be in format name="value"</param>
            public XmlAttributeItem(string xmlAttributeText)
            {
                Name = "";
                Value = "";
                char c;
                int readCount = 0, writeCount = 0;
                char[] buffer = new char[1024];
                for(readCount = 0; readCount<xmlAttributeText.Length; readCount++)
                {
                    c = xmlAttributeText[readCount];
                    if (c == '=') // get to the end of the attribute name.
                    {
                        Name = new string(buffer, 0, writeCount);
                        writeCount = 0;
                        break;
                    }
                    else
                    {
                        buffer[writeCount] = c;
                        writeCount++;
                    }
                }
                // finished reading the attribute name. move to the attribute value
                do
                {
                    readCount++;
                } while (xmlAttributeText[readCount] != '\"');
                readCount++;
                for (; readCount < xmlAttributeText.Length; readCount++)
                {
                    c = xmlAttributeText[readCount];
                    if (c == '\"') // get to the end of the attribute value
                    {
                        Value = new string(buffer, 0, writeCount);
                        break;
                    }
                    else
                    {
                        buffer[writeCount] = c;
                        writeCount++;
                    }
                }
            }
        }

        /// <summary>
        /// This struct contains two string arrays. header array contains the texts in the header. colRef array contains the corresponding column reference
        /// </summary>
        public struct HeaderWithColRef
        {
            public string[] header;
            public string[] colRef;
        }
    }

    /// <summary>
    /// (c) 2014 Vienna, Dietmar Schoder
    /// 
    /// Code Project Open License (CPOL) 1.02
    /// 
    /// Handles a "shared strings XML-file" in an Excel xlsx-file
    /// </summary>
    [Serializable()]
    [XmlType(Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    [XmlRoot("sst", Namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main")]
    public class SharedStringTable
    {
        [XmlAttribute]
        public string uniqueCount;
        [XmlAttribute]
        public string count;
        [XmlElement("si")]
        public SharedString[] si;

        public SharedStringTable()
        {
        }
    }
    public class SharedString
    {
        public string t;
        public override string ToString()
         => t;
    }
}
