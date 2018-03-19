package com.aspose.words.examples.mail_merge;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataTable;

import static com.aspose.words.examples.programming_documents.document.InsertDocumentIntoAnotherDocument.insertDocument;

/**
 * Sample code using Aspose.Words for Java.
 * Uses technique outlined here:
 * https://docs.aspose.com/display/wordsjava/How+to++Insert+a+Document+into+another+Document
 */
public class RepeatingSubDocuments
{
    //ExStart:
    private static final String dataDir = Utils.getSharedDataDir(ExecuteMailMergeWithRegions.class) + "SubDocuments/";

    public static void main(String[] args) throws Exception {

        // Construct based on Parent Document
        Document doc = new Document(dataDir + "RepeatingSubDocuments.ParentDocument.doc");

        // Pull 3 Orders from Northwind Database
        int[] orderIds = new int[] { 10406, 10412, 10444 };

        for ( int orderId : orderIds) {

            Document orderDoc = new Document(dataDir + "MailMerge.ExecuteWithRegions.doc");
            // Perform several mail merge operations populating only part of the document each time.
            // Use DataTable as a data source.
            // The table name property should be set to match the name of the region defined in the document.
            DataTable orderTable = getTestOrder(orderId);
            orderDoc.getMailMerge().executeWithRegions(orderTable);

            DataTable orderDetailsTable = getTestOrderDetails(orderId, "ExtendedPrice DESC");
            orderDoc.getMailMerge().executeWithRegions(orderDetailsTable);

            insertSubDocumentAtBookmark(doc, orderDoc, "SubDocumentBookmark");
        }

//        subDoc.save(dataDir + "CasePrint SubDocument out.doc");
        //insertSubDocumentAtBookmark(doc, subDoc, "SubDocumentBookmark");


        doc.save(dataDir + "RepeatingSubDocuments_Out.doc");
    }

    private static DataTable getTestOrder(int orderId) throws Exception {
        java.sql.ResultSet resultSet = executeDataTable(java.text.MessageFormat.format("SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", Integer.toString(orderId)));

        return new DataTable(resultSet, "Orders");
    }

    private static DataTable getTestOrderDetails(int orderId, String orderBy) throws Exception {
        StringBuilder builder = new StringBuilder();

        builder.append(java.text.MessageFormat.format("SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0}", Integer.toString(orderId)));

        if ((orderBy != null) && (orderBy.length() > 0)) {
            builder.append(" ORDER BY ");
            builder.append(orderBy);
        }

        java.sql.ResultSet resultSet = executeDataTable(builder.toString());
        return new DataTable(resultSet, "OrderDetails");
    }

    /**
     * Utility function that creates a connection, command, executes the command
     * and return the result in a DataTable.
     */
    private static java.sql.ResultSet executeDataTable(String commandText) throws Exception {
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String connString = "jdbc:ucanaccess://" + dataDir + "Northwind.mdb";

        // From Wikipedia: The Sun driver has a known issue with character encoding and Microsoft Access databases.
        // Microsoft Access may use an encoding that is not correctly translated by the driver, leading to the replacement
        // in strings of, for example, accented characters by question marks.
        //
        // In this case I have to set CP1252 for the European characters to come through in the data values.
        java.util.Properties props = new java.util.Properties();
        props.put("charSet", "Cp1252");

        // DSN-less DB connection.
        java.sql.Connection conn = java.sql.DriverManager.getConnection(connString, props);

        // Create and execute a command.
        java.sql.Statement statement = conn.createStatement();
        return statement.executeQuery(commandText);
    }
    //ExEnd:

    public static void insertSubDocumentAtBookmark(Document mainDoc, Document subDoc, String bookmarkName) throws Exception {
        Bookmark bookmark = mainDoc.getRange().getBookmarks().get(bookmarkName);
        insertDocument(bookmark.getBookmarkStart().getParentNode(), subDoc);
    }

}
