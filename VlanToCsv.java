package com.ibm.commops;

import java.io.File;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Iterator;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Row;
import org.odftoolkit.simple.table.Table;

public class VlanToCsv {

    private SpreadsheetDocument workBook = null;

    private Table sheet = null;

    private Row row = null;

    private Cell cell;

//    public void createWorkbook( String inputFileName ) throws Exception {
//        try {
//            workBook = SpreadsheetDocument.loadDocument( inputFileName );
//        } catch ( Exception e ) {
//            throw new Exception( e.getMessage() );
//        }
//
//    }

//    public ArrayList<String> getRow( String targetSheet, String firstItemInRow ) throws Exception {
//        // populate the Platform sheet object with the selected row
//
//        ArrayList<String> rowList = null;
//
//        sheet = workBook.getTableByName( targetSheet );
//        // iterate over rows
//        Iterator<Row> rowIt = sheet.getRowIterator();
//        while ( rowIt.hasNext() ) {
//            row = rowIt.next();
//            cell = row.getCellByIndex( 0 );
//            String cellContent = cell.getStringValue();
//            if ( cellContent.compareTo( firstItemInRow ) == 0 ) {
//                rowList = convertRowToList( row );
//                break;
//            }
//        }
//
//        return rowList;
//    }

    public ArrayList getAllRows( String sheetName, boolean skipFirstRow ) throws Exception {
        /*
         * A more generic implementation of the individual getAllSheetRows methods. Gets all rows in the sheet
         */
        ArrayList<ArrayList<String>> allRows = new ArrayList<ArrayList<String>>();

        sheet = workBook.getTableByName( sheetName );

        if ( sheet == null ) {
            throw new Exception( "The worksheet named : " + sheetName + " doesn't exist in the workbook" );
        }

        int sheetRowCount = getRowCount();
        for ( int i = 0; i < sheetRowCount; i++ ) {
            row = sheet.getRowByIndex( i );
            if ( skipFirstRow == true ) {
                skipFirstRow = false;
                // save the content for tests that do a callback for the title
                ArrayList<String> rowContent = convertRowToList( row );
                continue;

            }

            // found the targetRow now add the content to an arraylist
            ArrayList<String> rowContent = convertRowToList( row );
            if ( isEmptyRow( rowContent ) )
                continue;

            allRows.add( rowContent );

        }

        return allRows;
    }

    protected boolean isEmptyRow( ArrayList<String> rowContent ) {
        if ( rowContent.isEmpty() )
            return true;

        boolean onlySpace = true;
        Iterator<String> it = rowContent.iterator();
        int index = 0;
        while ( it.hasNext() ) {
            String content = it.next();

            // return true when the first cell is empty.
            if ( ( content.isEmpty() || content.equals( " " ) ) && index == 0 )
                return true;

            // return false when anything found
            if ( !content.isEmpty() && !content.equals( " " ) )
                return false;

            index++;
        }

        return true;
    }

    private ArrayList<String> convertRowToList( Row row ) {
        // Iterate over the cells in the row and create an Arraylist

        ArrayList<String> rowList = new ArrayList<String>();

        int minColIndex = 0;
        int maxColIndex = getColumnCount();

        for ( int colIndex = minColIndex; colIndex < maxColIndex; colIndex++ ) {

            cell = row.getCellByIndex( colIndex );
            String value = null;
            if ( cell == null )
                value = " ";
            else {
                if ( cell.getValueType() != null && cell.getValueType() == "boolean" ) {
                    if ( cell.getBooleanValue() )
                        value = "TRUE";
                    else
                        value = "FALSE";
                } else if ( cell.getValueType() != null && cell.getValueType() == "float" ) {
                    value = Double.toString( cell.getDoubleValue() );

                } else
                    value = cell.getStringValue();
            }

            rowList.add( value );
        }
        return rowList;

    }

    private int getColumnCount() {
        // The cell count is always the header count of the current sheet
        // This had to be created for the Symphony implementation to count the columns in the sheet based on which
        // ones have headers. Table.getColumnHeaderCount(), Table.getColumnCount() and Row.getCellCount() all don't work
        // in simple-odf 0.6.6
        // int count = 0;
        // if ( sheet == null )
        // return 0;
        //
        // // get the header row
        // Row row = sheet.getRowByIndex( 0 );
        //
        // for ( int i = 0; i < 200; i++ ) { // hardcoded 200 which will be the maxiumum amount column headers allowed.
        // // Don't think a sheet would ever have that much
        // String cellContent = row.getCellByIndex( i ).getStringValue();
        // if ( cellContent == null || cellContent.isEmpty() ) { // if no header name in the cell then that is the last
        // // column
        // count = i;
        // break;
        // }
        // }

        return 7;
    }

    private int getRowCount() {

        int count = 0;
        if ( sheet == null )
            return 0;

        // iterate over rows, while checking for "Test Name" cell. If it is empty, we're at the last row
        Iterator<Row> rowIt = sheet.getRowIterator();
        while ( rowIt.hasNext() ) {
            Row currRow = rowIt.next();
            // check the Test Name cell to see if it is empty
            if ( !currRow.getCellByIndex( 0 ).getStringValue().isEmpty() )
                count++;
            else
                break;
        }

        return count;

    }

    private static char[] filter( char[] inChars ) {
        String inStr = String.valueOf( inChars );

        inStr = inStr.replace( "WHRC, Teva", "WHRC/Teva" );
        inStr = inStr.replace( "DMZ   Internal", "DMZ Internal" );

        // remove commas
        inStr = inStr.replace( ", ", "/" );

        return inStr.toCharArray();
    }

    private static char[] toAscii( String inString ) {
        char[] inChars = inString.toCharArray();
        char[] outChars = new char[inChars.length];
        for ( int i = 0; i < inChars.length; i++ ) {
            if ( inChars[i] > 127 ) {
                // TODO: getting some odd characters so setting them to "-". See if this becomes an issue
                outChars[i] = 45;
            } else
                outChars[i] = inChars[i];
        }

        return outChars;
    }

    public static void main( String[] args ) {

        VlanToCsv reader = new VlanToCsv();
        try {
            reader.createWorkbook( "VLAN_Information.ods" );
            // reader.getRow( "Solution", "Name" );
            ArrayList<ArrayList<String>> out = reader.getAllRows( "Solution", true );
            ArrayList<ArrayList<String>> out2 = reader.getAllRows( "Production", true );

            for ( ArrayList<String> out3 : out2 ) {
                out.add( out3 );
            }

            File outFile = new File( "solutions.csv" );
            if ( !outFile.exists() )
                outFile.createNewFile();
            PrintWriter output = new PrintWriter( outFile, "windows-1252" );
            StringBuffer buff = new StringBuffer();

            output.println( "Environment,Parent Solution,Solution,Purpose,VLAN,Location" );
            for ( int i = 0; i < out.size(); i++ ) {
                char[] val = toAscii( out.get( i ).get( 5 ) );

                if ( ( val.length == 1 && (int) val[0] == 32 ) || val.length == 0 )
                    continue;

                output.print( filter( toAscii( out.get( i ).get( 4 ) ) ) );
                output.print( "," );
                output.print( parseParent( filter( toAscii( out.get( i ).get( 5 ) ) ) ) );
                output.print( "," );
                output.print( filter( toAscii( out.get( i ).get( 5 ) ) ) );
                output.print( "," );
                output.print( filter( toAscii( out.get( i ).get( 6 ) ) ) );
                output.print( "," );
                output.print( filter( toAscii( out.get( i ).get( 1 ) ) ) );
                output.print( "," );
                output.println( filter( toAscii( out.get( i ).get( 3 ) ) ) );

            }
            output.close();
        } catch ( Exception e ) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    private static char[] parseParent( char[] inChars ) {

        String inStr = String.valueOf( inChars );
        String[] outStr = inStr.split( "-" );
        if ( outStr.length > 1 )
            return outStr[0].toCharArray();

        outStr = inStr.split( " " );
        if ( outStr.length > 1 )
            return outStr[0].toCharArray();

        return inStr.toCharArray();
    }

}
