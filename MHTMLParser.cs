using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;
 
/// <summary>
/// HTMLParser is an object that can decode mhtml into ASCII text.
/// Using getHTMLText() will generate static HTML with inline images. 
/// </summary>
public class MHTMLParser
{
    const string BOUNDARY = "boundary";
    const string CHAR_SET = "charset";
    const string CONTENT_TYPE = "Content-Type";
    const string CONTENT_TRANSFER_ENCODING = "Content-Transfer-Encoding";
    const string CONTENT_LOCATION = "Content-Location";
    const string FILE_NAME = "filename=";
 
    private string mhtmlString; // the string we want to decode
    private string log; // log file
    public bool decodeImageData; //decode images?
 
    /*
     * Results of Conversion
     * This is split into a string[3] for each part
     * string[0] is the content type
     * string[1] is the content name
     * string[2] is the converted data
     */
    public List<string[]> dataset;
 
    /*
     * Default Constructor
     */
    public MHTMLParser()
    {
        this.dataset = new List<string[]>(); //Init dataset
        this.log += "Initialized dataset.\n";
        this.decodeImageData = false; //Set default for decoding images
    }
 
    /*
     * Init with contents of string 
     */
    public MHTMLParser(string mhtml)
        : this()
    {
        setMHTMLString(mhtml);
    }
    /*
     * Init with contents of string, and decoding option
     */
    public MHTMLParser(string mhtml, bool decodeImages)
        : this(mhtml)
    {
        this.decodeImageData = decodeImages;
    }
    /*
     * Set the mhtml string we want to decode
     */
    public void setMHTMLString(string mhtml)
    {
        try
        {
            if (mhtml == null) throw new Exception("The mhtml string is null"); //Early Exit
            this.mhtmlString = mhtml; //Set String
            this.log += "Set mhtml string.\n";
        }
        catch (Exception e)
        {
            this.log += e.Message;
            this.log += e.StackTrace;
        }
    }
    /*
     * Decompress Archive From String
     */
    public List<string[]> decompressString()
    {
        // init Prerequisites
        StringReader reader = null;
        string type = "";
        string encoding = "";
        string location = "";
        string filename = "";
        string charset = "utf-8";
        StringBuilder buffer = null;
        this.log += "Starting decompression \n";
 
 
        try
        {
            reader = new StringReader(this.mhtmlString); //Start reading the string
 
            String boundary = getBoundary(reader); // Get the boundary code
            if (boundary == null) throw new Exception("Failed to find string 'boundary'");
            this.log += "Found boundary.\n";
 
            //Loop through each line in the string
            string line = null;
            while ((line = reader.ReadLine()) != null)
            {
                string temp = line.Trim();
                if (temp.Contains(boundary)) //Check if this is a new section
                {
                    if (buffer != null) //If this is a new section and the buffer is full, write to dataset
                    {
                        string[] data = new string[3];
                        data[0] = type;
                        data[1] = filename;
                        data[2] = writeBufferContent(buffer, encoding, charset, type, this.decodeImageData);
                        this.dataset.Add(data);
                        buffer = null;
                        this.log += "Wrote Buffer Content and reset buffer.\n";
                    }
                    buffer = new StringBuilder();
                }
                else if (temp.StartsWith(CONTENT_TYPE))
                {
                    type = getAttribute(temp);
                    this.log += "Got content type.\n";
                }
                else if (temp.StartsWith(CHAR_SET))
                {
                    charset = getCharSet(temp);
                    this.log += "Got charset.\n";
                }
                else if (temp.StartsWith(CONTENT_TRANSFER_ENCODING))
                {
                    encoding = getAttribute(temp);
                    this.log += "Got encoding (" + encoding + ").\n";
                }
                else if (temp.StartsWith(CONTENT_LOCATION))
                {
                    location = temp.Substring(temp.IndexOf(":") + 1).Trim();
                    this.log += "Got location.\n";
                }
                else if (temp.StartsWith(FILE_NAME))
                {
                    char c = '"';
                    filename = temp.Substring(temp.IndexOf(c.ToString()) + 1, temp.LastIndexOf(c.ToString()) - temp.IndexOf(c.ToString()) - 1);
                }
                else if (temp.StartsWith("Content-ID") || temp.StartsWith("Content-Disposition") || temp.StartsWith("name=") || temp.Length == 1)
                {
                    //We don't need this stuff; Skip lines
                }
                else
                {
                    if (buffer != null)
                    {
                        buffer.Append(line + "\n");
                    }
                }
            }
        }
        finally
        {
            if (null != reader)
                reader.Close();
            this.log += "Closed Reader.\n";
        }
        return this.dataset; //Return Results
    }
    private string writeBufferContent(StringBuilder buffer, string encoding, string charset, string type, bool decodeImages)
    {
        this.log += "Start writing buffer contents.\n";
 
        //Detect if this is an image and if we want to decode it
        if (type.Contains("image"))
        {
            this.log += "Image Data Detected.\n";
            if (!decodeImages)
            {
                this.log += "Skipping image decode.\n";
                return buffer.ToString();
            }
        }
 
        // base64 Decoding
        if (encoding.ToLower().Equals("base64"))
        {
            try
            {
                this.log += "base64 encoding detected.\n";
                this.log += "Got base64 decoded string.\n";
                return decodeFromBase64(buffer.ToString());
            }
            catch (Exception e)
            {
                this.log += e.Message + "\n";
                this.log += e.StackTrace + "\n";
                this.log += "Data not Decoded.\n";
                return buffer.ToString();
            }
        }
        //quoted-printable decoding
        else if (encoding.ToLower().Equals("quoted-printable"))
        {
            this.log += "Quoted-Prinatble string detected.\n";
            return getQuotedPrintableString(buffer.ToString());
        }
        else
        {
            this.log += "Unknown Encoding.\n";
            return buffer.ToString();
        }
    }
    /*
     * Take base64 string, get bytes and convert to ascii string
     */
    static public string decodeFromBase64(string encodedData)
    {
        byte[] encodedDataAsBytes
            = System.Convert.FromBase64String(encodedData);
        string returnValue =
           System.Text.ASCIIEncoding.ASCII.GetString(encodedDataAsBytes);
        return returnValue;
    }
    /*
     * Get decoded quoted printable string
     */
    public string getQuotedPrintableString(string mimeString)
    {
        try
        {
            throw new Exception("Quoted-Printable is not supported.");
        }
        catch (Exception e)
        {
            this.log += e.Message + "\n";
            this.log += e.StackTrace + "\n";
            this.log += "Data not Decoded.\n";
            return mimeString;
        }
    }
    /*
     * Finds boundary used to break code into multiple parts
     */
    private string getBoundary(StringReader reader)
    {
        string line = null;
 
        while ((line = reader.ReadLine()) != null)
        {
            line = line.Trim();
            //If the line starts with BOUNDARY, lets grab everything in quotes and return it
            if (line.StartsWith(BOUNDARY))
            {
                char c = '"';
                int a = line.IndexOf(c.ToString());
                int b = line.LastIndexOf(c.ToString());
                return line.Substring(line.IndexOf(c.ToString()) + 1, line.LastIndexOf(c.ToString()) - line.IndexOf(c.ToString()) - 1);
            }
        }
        return null;
    }
    /*
     * Grabs charset from a line 
     */
    private string getCharSet(String temp)
    {
        string t = temp.Split('=')[1].Trim();
        return t.Substring(1, t.Length - 1);
    }
    /*
     * split a line on ": "
     */
    private string getAttribute(String line)
    {
        string str = ": ";
        return line.Substring(line.IndexOf(str) + str.Length, line.Length - (line.IndexOf(str) + str.Length)).Replace(";", "");
    }
    /*
     * Get an html page from the mhtml. Embeds images as base64 data
     */
    public string getHTMLText()
    {
        if (this.decodeImageData) throw new Exception("Turn off image decoding for valid html output.");
        List<string[]> data = this.decompressString();
        string body = "";
        //First, lets write all non-images to mail body
        //Then go back and add images in 
        for (int i = 0; i < 2; i++)
        {
            foreach (string[] strArray in data)
            {
                if (i == 0)
                {
                    if (strArray[0].Equals("text/html"))
                    {
                        body += strArray[2];
                        this.log += "Writing HTML Text\n";
                    }
                }
                else if (i == 1)
                {
                    if (strArray[0].Contains("image"))
                    {
                        body = body.Replace("cid:" + System.Environment.NewLine, "cid:").Replace("cid:" + strArray[1], "data:" + strArray[0] + ";base64," + strArray[2]);
                        this.log += "Overwriting HTML with image: " + strArray[1] + "\n";
                    }
                }
            }
        }
        return body;
    }
    /*
     *  Get the log from the decoding process
     */
    public string getLog()
    {
        return this.log;
    }
}
