MHTML-to-HTML-Decoding-in-C-
============================

MHTML (short for MIME HTML) is a web archive that stores a web page’s HTML and (normally remote) resources in one file. It is composed in a manner similar to an HTML email, using the content-type ‘multipart-related’. The data is split into parts and base64 encoded.

Although this code will decode .mht and .mhtml files, in it’s current state it will only decode the base64 content-transfer encoding. It has been tested on .mhtml files exported from SQL Server Reporting Service (SSRS). It features it’s own logging and a way return valid HTML (with images)

The return of the decompression value is a List<string[]>. Each List element is a section of the MHTML, and the contents of each List element is as follows:
string[0] is the Content-Type
string[1] is the Content-Name
string[2] is the converted data

Using the getHTMLText() method will return the full HTML and will use the cid:’s to insert the base64 image data (valid in newer browsers).



And here is how to use it

string mhtml = "This is your MHTML string"; // Make sure the string is in UTF-8 encoding
MHTMLParser parser = new MHTMLParser(mhtml);
string html = parser.getHTMLText(); // This is the converted HTML
