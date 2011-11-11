/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package cm.google;

import com.Ostermiller.util.CSVParser;
import com.google.gdata.client.GoogleAuthTokenFactory.UserToken;
import com.google.gdata.client.docs.DocsService;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.MediaContent;
import com.google.gdata.data.media.MediaSource;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringWriter;
import java.io.UnsupportedEncodingException;
import java.net.MalformedURLException;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

/**
 *
 * @author codemitte
 */
public class SpreadsheetDumper {

    public static String APP_NAME = "codemitte-GoogleSpreadsheetDumper-0.1";
    public DocsService client;
    public SpreadsheetService spreadClient;

    public String export(String[] args) {
        String documentId = args[0];
        String targetLoc = args[1];

        System.err.println("# Dumping " + documentId + " to " + targetLoc);

        String username = args.length > 2 ? args[2] : null;
        String password = args.length > 3 ? args[3] : null;

        java.io.Console cons;
        char[] passwd = null;
        if ((cons = System.console()) != null
                && (username != null || (username = cons.readLine("[%s] ", "Username:")) != null)
                && (password != null || (passwd = cons.readPassword("[%s] ", "Password:")) != null)) {
            if (passwd != null) {
                password = String.copyValueOf(passwd);
            }
        }

        System.err.println("# User " + username + " identified by " + password);


        // START DOWNLOADING THE SPREADSHEET
        client = new DocsService(APP_NAME);

        spreadClient = new SpreadsheetService(APP_NAME);
        try {
            spreadClient.setUserCredentials(username, password);

            // Substitute the spreadsheets token for the docs token
            UserToken spreadsheetsToken = (UserToken) spreadClient.getAuthTokenFactory().getAuthToken();
            client.setUserToken(spreadsheetsToken.getValue());


        } catch (AuthenticationException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);

            System.exit(1);
        }

        Map<String, List<String>> map;
        try {
            map = getUrlParameters(documentId);

            documentId = map.get("key").get(0);
        } catch (UnsupportedEncodingException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        }


        System.err.println("# Found documentId " + documentId);
        try {
            downloadSpreadsheet(documentId, targetLoc);
        } catch (IOException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ServiceException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        }

        return targetLoc;
    }

    public void downloadSpreadsheet(String docId, String filepath)
            throws IOException, MalformedURLException, ServiceException {
        String fileExtension = filepath.substring(filepath.lastIndexOf(".") + 1);
        String exportUrl = "https://spreadsheets.google.com/feeds/download/spreadsheets"
                + "/Export?key=" + docId + "&exportFormat=" + fileExtension;

        // If exporting to .csv or .tsv, add the gid parameter to specify which sheet to export
        if (fileExtension.equals("csv") || fileExtension.equals("tsv")) {
            exportUrl += "&gid=0";  // gid=0 will download only the first sheet
        }

        downloadFile(exportUrl, filepath);
    }

    public void downloadFile(String exportUrl, String filepath)
            throws IOException, MalformedURLException, ServiceException {
        System.out.println("Exporting document from: " + exportUrl);

        MediaContent mc = new MediaContent();
        mc.setUri(exportUrl);
        MediaSource ms = client.getMedia(mc);

        InputStream inStream = null;
        FileOutputStream outStream = null;

        try {
            inStream = ms.getInputStream();
            outStream = new FileOutputStream(filepath);

            int c;
            while ((c = inStream.read()) != -1) {
                outStream.write(c);
            }
        } finally {
            if (inStream != null) {
                inStream.close();
            }
            if (outStream != null) {
                outStream.flush();
                outStream.close();
            }
        }
    }

    public static Map<String, List<String>> getUrlParameters(String url)
            throws UnsupportedEncodingException {
        Map<String, List<String>> params = new HashMap<String, List<String>>();
        String[] urlParts = url.split("\\?");
        if (urlParts.length > 1) {
            String query = urlParts[1];
            for (String param : query.split("&")) {
                String pair[] = param.split("=");
                String key = URLDecoder.decode(pair[0], "UTF-8");
                String value = URLDecoder.decode(pair[1], "UTF-8");
                List<String> values = params.get(key);
                if (values == null) {
                    values = new ArrayList<String>();
                    params.put(key, values);
                }
                values.add(value);
            }
        }
        return params;
    }

    public static void printPackageXml(String fileName) throws ParserConfigurationException, TransformerConfigurationException, TransformerException, FileNotFoundException, IOException {
        //We need a Document
        DocumentBuilderFactory dbfac = DocumentBuilderFactory.newInstance();
        DocumentBuilder docBuilder = dbfac.newDocumentBuilder();
        Document doc = docBuilder.newDocument();

        ////////////////////////
        //Creating the XML tree

        //create the root element and add it to the document
        org.w3c.dom.Element root = doc.createElement("Package");
        doc.appendChild(root);

        // build xml - after reading csv
        CSVParser in = new CSVParser(new FileInputStream(new File(fileName)));
        in.changeDelimiter(',');

        HashMap<String, ArrayList<String>> data = new HashMap<String, ArrayList<String>>();

        String[] line = in.getLine();
        
        while (line != null) {
            String entity = line[0];
            String type = line[1];

            if (in.getLastLineNumber()  <= 1)
            {
                line = in.getLine();
                continue;
            }
            
            // mapping for lazy users
            if ("class".equals(type.toLowerCase()) || ".".equals(type)) {
                type = "ApexClass";
            } else if ("trigger".equals(type.toLowerCase())) {
                type = "ApexTrigger";
            } else if ("page".equals(type.toLowerCase())) {
                type = "ApexPage";
            } else if ("trigger".equals(type.toLowerCase())) {
                type = "ApexTrigger";
            } else if ("field".equals(type.toLowerCase())) {
                type = "CustomField";
            } else if ("label".equals(type.toLowerCase())) {
                type = "CustomLabels";
            }

            if (!data.containsKey(type)) {
                data.put(type, new ArrayList<String>());
            }
            
            if (!entity.startsWith("."))
            {
                data.get(type).add(entity);
            }

            line = in.getLine();
        }


        //Output the XML
        for (String type : data.keySet())
        {
            Element types = doc.createElement("types");
            
            Element typeName = doc.createElement("typeName");
            typeName.appendChild(doc.createTextNode(type));
            types.appendChild(typeName);
            
            for (String member : data.get(type))
            {
                Element members = doc.createElement("members");
                members.appendChild(doc.createTextNode(member));
                types.appendChild(members);
            }
            
            root.appendChild(types);
        }


        //set up a transformer
        TransformerFactory transfac = TransformerFactory.newInstance();
        Transformer trans = transfac.newTransformer();
        trans.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
        trans.setOutputProperty(OutputKeys.INDENT, "yes");

        //create string from xml tree
        StringWriter sw = new StringWriter();
        StreamResult result = new StreamResult(sw);
        DOMSource source = new DOMSource(doc);
        trans.transform(source, result);
        String xmlString = sw.toString();

        //print xml
        System.out.println(xmlString);
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        if (args.length < 2) {
            System.err.println("Usage: java SpreadsheetDumper <DocumentId> <targetLocation> [<username> [<password>]]");
            System.exit(1);
        }
        try {
            printPackageXml(new SpreadsheetDumper().export(args));
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerConfigurationException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (TransformerException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(SpreadsheetDumper.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
