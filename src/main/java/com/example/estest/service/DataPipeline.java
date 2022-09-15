package com.example.estest.service;


import com.fasterxml.jackson.databind.exc.InvalidFormatException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.*;
import java.nio.charset.StandardCharsets;
import java.text.NumberFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Objects;

public class DataPipeline {

    public static void main(String[] args) throws IOException {
        Sample sample = new Sample();


        for (int i = 3001 ; i <= 5500; i += 500) {
            sample.run(i, i + 499);
        }

    }

    public static class Sample {
        private final String defaultPath = "./data/";
        private final String csvPath = "./csv/";
        private final URL baseUrl;

        {
            try {
                baseUrl = new URL("http://kpat.kipris.or.kr/kpat/downa.do?method=download");
                File theDir = new File(defaultPath);
                File csvDir = new File(csvPath);
                if (!theDir.exists() | !csvDir.exists()) {
                    theDir.mkdirs();
                    csvDir.mkdirs();
                }
            } catch (MalformedURLException e) {
                throw new RuntimeException(e);
            }
        }

        /**
         * 바이너리데이터를 파일로 저장
         *
         * @param input:      response로 받은 binary data
         * @param output:     파일
         * @param bufferSize: bufferSize
         * @throws IOException
         */
        public static void copySource2Dest(InputStream input, OutputStream output, int bufferSize)
                throws IOException {
            byte[] buf = new byte[bufferSize];
            int n = input.read(buf);
            while (n >= 0) {
                output.write(buf, 0, n);
                n = input.read(buf);
            }
            output.flush();
        }


        public String fetchFileDownloadUrlByExpression(int from, int to) throws IOException {
            System.out.println("fetchFileDownloadUrlByExpression 시작!");
            HttpURLConnection conn = (HttpURLConnection) baseUrl.openConnection();
            HashMap<String, String> payload = makePayload(from, to);
            String postDataString = payloadToString(payload);
            byte[] postData = postDataString.getBytes(StandardCharsets.UTF_8);
            conn.setDoOutput(true);
            conn.setInstanceFollowRedirects(false);
            conn.setRequestMethod("POST");
            conn.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
            conn.setRequestProperty("charset", "utf-8");
            conn.setRequestProperty("Content-Length", Integer.toString(postData.length));
            conn.setUseCaches(false);
            try (DataOutputStream wr = new DataOutputStream(conn.getOutputStream())) {
                wr.write(postData);
                wr.flush();
            }
            int responseCode = conn.getResponseCode();
            System.out.println("POST Response Code :: " + responseCode);

            if (responseCode == HttpURLConnection.HTTP_OK) { //success
                BufferedReader in = new BufferedReader(new InputStreamReader(
                        conn.getInputStream()));
                String inputLine;
                StringBuffer response = new StringBuffer();
                while ((inputLine = in.readLine()) != null) {
                    response.append(inputLine);
                }
                in.close();
                // print result
                return response.toString();
            } else {
                System.out.println("POST request not worked");
                return "";
            }
        }

        /**
         * 실행! (expression을 받아 특허 엑셀을 다운로드 받음)
         *
         * @param from,to: 검색어
         * @return: 성공시:1 실패시:0
         */
        public int run(int from, int to) {
            System.out.println("시작!");
            try {
                String fileUrl = fetchFileDownloadUrlByExpression(from, to);
                if (Objects.equals(fileUrl, "-1")) {
                    System.out.println("Error 다시 시도");
                    return 0;
                }
                String filename = downloadFile(fileUrl);
                System.out.println(filename);
                if (filename == null) {
                    return 0;
                }
                testxls(filename, from, to);


            } catch (IOException e) {
                throw new RuntimeException(e);
            }
            return 1;
        }

        /**
         * 해당 url의 파일을 다운로드 받음
         *
         * @param fileUrl
         * @return if success: true else false
         * @throws IOException: IDK
         */
        String downloadFile(String fileUrl) throws IOException {
            System.out.println("다운로드 시작!");
            URL u = new URL(fileUrl);
            String fileName = fileUrl.substring(fileUrl.lastIndexOf('/') + 1, fileUrl.length());
            URLConnection conn = u.openConnection();
            InputStream in = conn.getInputStream();
            FileOutputStream out = new FileOutputStream(defaultPath + fileName);

            copySource2Dest(in, out, 1024);
            out.close();
            return fileName;
        }

        public void xlstocsv(String filename) throws IOException {
            // Creating a inputFile object with specific file path
            File inputFile = new File(defaultPath + filename + ".xls");

            // Creating a outputFile object to write excel data to csv
            File outputFile = new File(csvPath + filename + ".csv");

            // For storing data into CSV files
            StringBuffer data = new StringBuffer();

            try {
                // Creating input stream
                FileInputStream fis = new FileInputStream(inputFile);
                Workbook workbook = null;

                // Get the workbook object for Excel file based on file format
                if (inputFile.getName().endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else if (inputFile.getName().endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                } else {
                    fis.close();
                    throw new Exception("File not supported!");
                }

                // Get first sheet from the workbook
                Sheet sheet = workbook.getSheetAt(0);

                // Iterate through each rows from first sheet
                Iterator<Row> rowIterator = sheet.iterator();

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    // For each row, iterate through each columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {

                        Cell cell = cellIterator.next();

                        switch (cell.getCellType()) {
                            case BOOLEAN:
                                data.append(cell.getBooleanCellValue() + ",");
                                break;

                            case NUMERIC:
                                data.append(cell.getNumericCellValue() + ",");
                                break;
                            case STRING:
                                if(cell.getStringCellValue().contains(",")){
                                    // , " 처리
                                    if (cell.getStringCellValue().contains("\"")){
                                        data.append("\"").append(cell.getStringCellValue().replaceAll("\"", "\"\"")).append("\",");
                                    }
                                    // , 처리
                                    else{
                                        data.append("\""+cell.getStringCellValue()+"\""+",");
                                    }
                                }
                                // " 처리
                                else if(cell.getStringCellValue().contains("\"")) {
                                    data.append("\"").append(cell.getStringCellValue().replaceAll("\"", "\"\"")).append("\",");
                                }
                                else if (cell.getStringCellValue().contains("skip")) {
                                    continue;
                                }
                                else{
                                    data.append(cell.getStringCellValue() + ",");
                                }
                                break;

                            case BLANK:
                                data.append("" + ",");
                                break;

                            default:
                                data.append(cell + ",");
                        }
                    }
                    // appending new line after each row
                    data.append('\n');
                }

                FileOutputStream fos = new FileOutputStream(outputFile);
                fos.write(data.toString().getBytes());
                fos.close();

            } catch (Exception e) {
                e.printStackTrace();
            }

            System.out.println("Conversion of an Excel file to CSV file is done!");

        }

        /**
         * payload를 만들어서 hashmap으로 리턴
         */
        private HashMap<String, String> makePayload(int from, int to) {
            HashMap<String, String> map = new HashMap<>();
            map.put("bookMark", "");
            map.put("expression", "AD=[19450101~20220908]");
            map.put("prefixExpression", "");
            map.put("sortField", "");
            map.put("sortState", "");
            map.put("sortField1", "AD");
            map.put("sortState1", "DESC");
            map.put("sortField2", "");
            map.put("sortState2", "");
            map.put("collections", "");
            map.put("BOOKMARK", "");
            map.put("byBookMark", "0");
            map.put("rangeAll", "0");
            map.put("viewField", "AN,AD,TTL,APPLICANT,IPC,CPC,PN,OPN,RRN,RGST_STCD,ABS_TEXT,");
            map.put("inclDraw", "0");
            map.put("inclJudg", "0");
            map.put("inclReg", "0");
            map.put("inclAdmin", "0");
            map.put("fileType", "xls");
            map.put("currPageNo", "1");
            map.put("fromNum", String.valueOf(from));
            map.put("toNum", String.valueOf(to));
            map.put("searchInTrans", "N");
            map.put("piField", "");
            map.put("piValue", "");
            map.put("piSearchYN", "N");
            return map;
        }

        /**
         * payload를 받아 url-formencoded로 만듬
         *
         * @param payload
         * @throws UnsupportedEncodingException
         * @return: form-encoded PayloadData
         */
        private String payloadToString(HashMap<String, String> payload) throws UnsupportedEncodingException {
            StringBuilder result = new StringBuilder();
            boolean first = true;
            for (Map.Entry<String, String> entry : payload.entrySet()) {
                if (first)
                    first = false;
                else
                    result.append("&");
                result.append(URLEncoder.encode(entry.getKey(), "UTF-8"));
                result.append("=");
                result.append(URLEncoder.encode(entry.getValue(), "UTF-8"));
            }
            System.out.println(result);
            return result.toString();
        }

        public String testxls(String filename, int from, int to) throws IOException, InvalidFormatException {

            String patentdata = "patent" + from + "-" + to;
            String newfile = defaultPath + patentdata + ".xls";
            File oldfile = new File(defaultPath + filename);
            InputStreamReader isr = new InputStreamReader(System.in);
            FileInputStream file = new FileInputStream(defaultPath + filename);

            NumberFormat f = NumberFormat.getInstance();
            f.setGroupingUsed(false);

            try {
                HSSFWorkbook workbook = new HSSFWorkbook(file);

                HSSFSheet sheet = workbook.getSheetAt(0);
                for (int i = 0; i < 7; i++) {
                    if (sheet.getRow(i) == null) {
                        continue;
                    } else {
                        sheet.removeRow(sheet.getRow(i));
                    }
                }
                int rows = sheet.getPhysicalNumberOfRows();
                System.out.println(rows);
                for (int i = 0; i< rows+7; i ++) {
                    if (sheet.getRow(i) == null) {
                        continue;
                    }
                    else {
                        sheet.getRow(i).getCell(0).setCellValue("skip");
                    }

                }

                FileOutputStream fileOutputStream = new FileOutputStream(defaultPath + patentdata + ".xls");
                workbook.write(fileOutputStream);

                if (oldfile.exists()) {
                    oldfile.delete();
                }

                workbook.close();
                xlstocsv(patentdata);

            } catch (Exception e) {
                e.printStackTrace();
            }

            return newfile;

        }


    }
}