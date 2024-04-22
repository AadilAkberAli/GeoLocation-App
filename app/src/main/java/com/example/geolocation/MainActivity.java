package com.example.geolocation;

import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;

import android.Manifest;
import android.app.Activity;
import android.content.Context;
import android.content.pm.PackageManager;
import android.content.res.AssetManager;
import android.location.Address;
import android.location.Geocoder;
import android.os.Bundle;
import android.os.Environment;
import android.util.Log;
import android.webkit.MimeTypeMap;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class MainActivity extends AppCompatActivity {
    private TextView textView;
     double latitude = 0;
    double longitude = 0;
    int totalrows = 0;
    int getrowcount = 0;
    ArrayList<String> anyvaluerow = new ArrayList<>();
    ArrayList<String> header= new ArrayList<>();
    List<Address> addresses = null;
    Workbook workbook;
    Geocoder geocoder;
    List<Geolocationmodel> geolocationmodelslist = new ArrayList<>();
    private static Cell cell;
    private static Sheet sheet;
    private static String EXCEL_SHEET_NAME = "Sheet1";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        textView = findViewById(R.id.textview);
        geocoder = new Geocoder(this, Locale.getDefault());

        extractExcelContentByColumnIndex(2,3, this, "location.xls", "new.xls");
        
 // Here 1 represent max location result to returned, by documents it recommended 1 to 5
//
//        String address = addresses.get(0).getAddressLine(0); // If any additional address line present than only, check with max available address lines by getMaxAddressLineIndex()
//        String city = addresses.get(0).getLocality();
//        String state = addresses.get(0).getAdminArea();
//        String country = addresses.get(0).getCountryName();
//        String postalCode = addresses.get(0).getPostalCode();
//        String knownName = addresses.get(0).getFeatureName();
    }

    public void extractExcelContentByColumnIndex(Integer latcolumn, Integer longcolumn,  Activity context, String filename, String newfilename)
    {
         final int REQUEST_EXTERNAL_STORAGE = 1;
         String[] PERMISSIONS_STORAGE = {
                Manifest.permission.READ_EXTERNAL_STORAGE,
                Manifest.permission.WRITE_EXTERNAL_STORAGE,
                 Manifest.permission.MANAGE_EXTERNAL_STORAGE
        };

        int permission = ActivityCompat.checkSelfPermission(context, Manifest.permission.READ_EXTERNAL_STORAGE);

        if (permission != PackageManager.PERMISSION_GRANTED) {
            // We don't have permission so prompt the user
            ActivityCompat.requestPermissions(
                    context,
                    PERMISSIONS_STORAGE,
                    REQUEST_EXTERNAL_STORAGE
            );
        }
        else
        {
            AssetManager assetManager = getResources().getAssets();
            InputStream fileInputStream = null;
            HSSFWorkbook workbook = null;
            XSSFWorkbook xssfWorkbook = null;
            try {
                try {
                    fileInputStream = assetManager.open(filename);
                } catch (IOException e) {
                    e.printStackTrace();
                }
                String extension = MimeTypeMap.getFileExtensionFromUrl(filename);
                if(extension.equalsIgnoreCase("xls"))
                {
                    // Create a workbook using the File System
                    POIFSFileSystem myFileSystem = new POIFSFileSystem(fileInputStream);
                    workbook = new HSSFWorkbook(myFileSystem);
                    // Get the first sheet from workbook
                    HSSFSheet  sheet = workbook.getSheetAt(0);
                    int noOfColumns = sheet.getRow(0).getLastCellNum();
                    totalrows = sheet.getPhysicalNumberOfRows() - 1;
                    for(Row r : sheet)
                    {
                        ReadColumndata(r,latcolumn,longcolumn, noOfColumns);
                    }
                }
                else
                {
                    xssfWorkbook = new XSSFWorkbook(fileInputStream);
                    // Get the first sheet from workbook
                    XSSFSheet  sheet = xssfWorkbook.getSheetAt(0);
                    int noOfColumns = sheet.getRow(0).getLastCellNum();
                    totalrows = sheet.getPhysicalNumberOfRows() - 1;
                    for(Row r : sheet)
                    {
                        ReadColumndata(r,latcolumn,longcolumn, noOfColumns);
                    }
                }
                if(geolocationmodelslist.size() > 0)
                {
                    exportDataIntoWorkbook(context,newfilename,geolocationmodelslist);
                }
                else
                {
                    Toast.makeText(getApplicationContext(),"Some thing went wrong", Toast.LENGTH_SHORT).show();
                }
            } catch (IOException e) {
                e.printStackTrace();
                Log.e("Error", e.toString());
            }

        }

    }

    public void ReadColumndata(Row r, Integer latcolumn, Integer longcolumn, int noOfColumns)
    {
        anyvaluerow=new ArrayList<>();
        Cell c = r.getCell(latcolumn);
        Cell c1 = r.getCell(longcolumn);
        for(int i=0; i<noOfColumns ;i++)
        {
            if(i == latcolumn)
            {
                if(c != null) {

                    if(c.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        latitude = c.getNumericCellValue();
                    }
                    else
                    {
                        header.add(c.getStringCellValue());
                    }
                }
            }
            else if(i == longcolumn)
            {
                if(c1 != null) {
                    if(c1.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                        longitude = c1.getNumericCellValue();
                    }
                    else
                    {
                        header.add(c1.getStringCellValue());
                    }
                }
            }
            else
            {
                Cell cn = r.getCell(i);
                if(cn != null) {
                    if(cn.getColumnIndex() == i && cn.getRowIndex() == 0)
                    {
                        header.add(cn.getStringCellValue());
                    }
                    else
                    {
                        if(cn.getCellType() == Cell.CELL_TYPE_NUMERIC)
                        {
                            anyvaluerow.add(String.valueOf(cn.getNumericCellValue()));
                        }
                        else
                        {
                            anyvaluerow.add(cn.getStringCellValue());
                        }
                    }

                }
            }
        }

        if(latitude > 0)
        {
            Log.e("completeheader", header.toString());
            Log.e("anyrowvalue", anyvaluerow.toString());
            try {
                Thread.sleep(500);
                getlocation(latitude, longitude, anyvaluerow);
            } catch (InterruptedException | IOException e) {
                e.printStackTrace();
            }
        }
    }
    private void getlocation(double latitude, double longitude, List<String> anyvaluerow) throws IOException {
        addresses = geocoder.getFromLocation(latitude, longitude, 1);
        String address = addresses.get(0).getAddressLine(0);
        geolocationmodelslist.add(new Geolocationmodel(anyvaluerow, latitude, longitude, address));
    }

    public  boolean exportDataIntoWorkbook(Context context, String fileName,
                                                 List<Geolocationmodel> dataList) {
        boolean isWorkbookWrittenIntoStorage;
        // Check if available and not read only
        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.e("main", "Storage not available or read only");
            return false;
        }

        // Creating a New HSSF Workbook (.xls format)
        workbook = new HSSFWorkbook();

        setHeaderCellStyle();

        // Creating a New Sheet and Setting width for each column
        sheet = workbook.createSheet(EXCEL_SHEET_NAME);
        sheet.setColumnWidth(0, (15 * 400));
        sheet.setColumnWidth(1, (15 * 400));
        sheet.setColumnWidth(2, (15 * 400));
        sheet.setColumnWidth(3, (15 * 400));

        setHeaderRow();
        fillDataIntoExcel(dataList);
        isWorkbookWrittenIntoStorage = storeExcelInStorage(context, fileName);

        return isWorkbookWrittenIntoStorage;
    }

    /**
     * Checks if Storage is READ-ONLY
     *
     * @return boolean
     */
    private static boolean isExternalStorageReadOnly() {
        String externalStorageState = Environment.getExternalStorageState();
        return Environment.MEDIA_MOUNTED_READ_ONLY.equals(externalStorageState);
    }

    /**
     * Checks if Storage is Available
     *
     * @return boolean
     */
    private static boolean isExternalStorageAvailable() {
        String externalStorageState = Environment.getExternalStorageState();
        return Environment.MEDIA_MOUNTED.equals(externalStorageState);
    }

    /**
     * Setup header cell style
     */
    private  void setHeaderCellStyle() {
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(HSSFColor.AQUA.index);
        headerCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        headerCellStyle.setAlignment(CellStyle.ALIGN_CENTER);
    }

    /**
     * Setup Header Row
     */
    private  void setHeaderRow() {
        header.add("Geolocation");
        Row row = sheet.createRow(0);
        CellStyle cellStyle = workbook.createCellStyle();

        for(int i=0; i<header.size(); i++)
        {
            cell = row.createCell(i);
            cell.setCellValue(header.get(i));
            cell.setCellStyle(cellStyle);
        }

    }

    /**
     * Fills Data into Excel Sheet
     * <p>
     * NOTE: Set row index as i+1 since 0th index belongs to header row
     *
     * @param dataList - List containing data to be filled into excel
     */
    private  void fillDataIntoExcel(List<Geolocationmodel> dataList) {
        for (int i = 0; i < dataList.size(); i++) {
            // Create a New Row for every new entry in list
            Row rowData = sheet.createRow(i + 1);

            // Create Cells for each row
            for(int col=0 ; col < dataList.get(i).getAnyvaluerow().size() ; col++)
            {

                cell = rowData.createCell(col);
                cell.setCellValue(geolocationmodelslist.get(getrowcount).getAnyvaluerow().get(col));
            }
            if(getrowcount < geolocationmodelslist.size())
            {
                getrowcount ++;
            }
            int latcolumn =  dataList.get(i).getAnyvaluerow().size() ;
            cell = rowData.createCell(latcolumn);
            cell.setCellValue(dataList.get(i).getLatitude());

            int longcolumn = latcolumn + 1;
            cell = rowData.createCell(longcolumn);
            cell.setCellValue(dataList.get(i).getLongitude());

            int address = longcolumn + 1;
            cell = rowData.createCell(address);
            cell.setCellValue(dataList.get(i).getAddress());

            textView.append(dataList.get(i).getAddress());
            textView.append("\n");
        }
    }

    /**
     * Store Excel Workbook in external storage
     *
     * @param context  - application context
     * @param fileName - name of workbook which will be stored in device
     * @return boolean - returns state whether workbook is written into storage or not
     */
    private  boolean storeExcelInStorage(Context context, String fileName) {
        boolean isSuccess;
        File file = new File(context.getExternalFilesDir(null), fileName);
        FileOutputStream fileOutputStream = null;

        try {
            fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            Log.e("file", "Writing file" + file);
            Toast.makeText(getApplicationContext(),"File save successfully", Toast.LENGTH_SHORT).show();
            isSuccess = true;
        } catch (IOException e) {

            Log.e("exception", "Error writing Exception: ", e);
            isSuccess = false;
        } catch (Exception e) {
            Log.e("failed", "Failed to save file due to Exception: ", e);
            isSuccess = false;
        } finally {
            try {
                if (null != fileOutputStream) {
                    fileOutputStream.close();
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return isSuccess;
    }
}

