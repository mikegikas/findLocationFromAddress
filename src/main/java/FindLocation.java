import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.util.*;


public class FindLocation {
    public static void main(String[] args) throws Exception {
        String googleUrl = "https://maps.googleapis.com/maps/api/geocode/json?address=";

        // Add your own google api key
        String apiKey = "";

        String inputFilePath = "./data/address.xlsx";

        String outputFilePath = "./data/LocationPointsResults.xlsx";

        DataFormatter formatter = new DataFormatter(Locale.US);

        ArrayList<String> address = new ArrayList<>();

        Map<Integer, Object[]> locationInfo = new TreeMap<>();

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet spreadsheet = workbook.createSheet("Location Info");

        XSSFRow writeRow;

        locationInfo.put(1, new Object[]{
                "Address", "Latitude", "Longitude", "Formatted Address"});

        try {
            Workbook wb = WorkbookFactory.create(new FileInputStream(inputFilePath));

            Sheet sheet = wb.getSheetAt(0);

            int rowsCount = sheet.getLastRowNum();

            for (int i = 1; i <= rowsCount; i++) {
                StringBuilder formatAddress = new StringBuilder();

                Row row = sheet.getRow(i);

                int colCounts = row.getLastCellNum();
                for (int j = 0; j < colCounts; j++) {
                    if (j == 0) {
                        formatAddress.append(formatter.formatCellValue(row.getCell(j)));
                    } else {
                        formatAddress.append(" ").append(formatter.formatCellValue(row.getCell(j)));
                    }
                }
                System.out.println("Actual Address After Concat Cols: " + formatAddress);
                address.add(formatAddress.toString());
            }
        } catch (Exception e) {
            System.out.println(e);
        }

        int rowInfo = 1;
        for (String word : address) {
            rowInfo++;

            StringBuilder actualUrl = new StringBuilder();

            String encodeWord = URLEncoder.encode(word, "UTF-8");

            actualUrl.append(googleUrl).append(encodeWord).append("&key=").append(apiKey);

            JSONObject jsonObject = new JSONObject(getUrlConnection(new URL(actualUrl.toString())));

            JSONArray jsonArray = jsonObject.getJSONArray("results");

            jsonObject = jsonArray.getJSONObject(0);

            String formatted_address = jsonObject.getString("formatted_address");

            JSONObject geometry = jsonObject.getJSONObject("geometry");

            JSONObject location = geometry.getJSONObject("location");

            Double latitude = location.getDouble("lat");

            Double longitude = location.getDouble("lng");

            locationInfo.put(rowInfo, new Object[]{
                    word, latitude.toString(), longitude.toString(), formatted_address});

            System.out.println(actualUrl);
            System.out.println("The Location you searched: " + word);
            System.out.println(formatted_address);
            System.out.println(latitude);
            System.out.println(longitude);
        }

        Set<Integer> keyId = locationInfo.keySet();
        int rowId = 0;

        for (Integer key : keyId) {
            writeRow = spreadsheet.createRow(rowId++);
            Object[] objectArr = locationInfo.get(key);
            int cellId = 0;

            for (Object obj : objectArr) {
                Cell cell = writeRow.createCell(cellId++);
                cell.setCellValue((String) obj);
            }
        }

        FileOutputStream out = new FileOutputStream(
                new File(outputFilePath));

        workbook.write(out);
        out.close();
        System.out.println("LocationPointsResults.xlsx written successfully");
    }

    public static String getUrlConnection(URL inputUrl) throws Exception {
        StringBuilder content = new StringBuilder();
        HttpURLConnection httpURLConnection = (HttpURLConnection) inputUrl.openConnection();
        httpURLConnection.setRequestMethod("GET");
        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(httpURLConnection.getInputStream()));
        String inputLine;
        while ((inputLine = bufferedReader.readLine()) != null) {
            content.append(inputLine);
        }
        bufferedReader.close();
        return content.toString();
    }
}
