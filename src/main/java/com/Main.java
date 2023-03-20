package com;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {

    private static int count = 0;

    public static String excel() throws IOException {

        List<String> list = new ArrayList<>();

        try (Workbook workbook = WorkbookFactory.create(new File("src/main/java/com/Result_1.xlsx"))) {
            int sheet = workbook.getNumberOfSheets();
            for (int i = 0; i < sheet; i++) {
                Sheet s = workbook.getSheetAt(i);
                Iterator<Row> rowIterator = s.rowIterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();

                    if (count > 0 && count < 420) {
                        String id = getId(row.getCell(0).toString());
                        String firstName = getFirstName(row.getCell(1).toString());
                        String lastName = getLastName(row.getCell(2).toString());
                        list.add(getQuery(id, firstName, lastName));
                    }
                    count++;
                }
            }
        }
        return toJSON(list);
    }

    private static String getQuery(String id, String firstName, String lastName) {
        String query = "UPDATE user.profile SET first_name = '%s', last_name = '%s', last_modified_date = NOW() WHERE id = %s;";
        return String.format(query, firstName, lastName, id);
    }

    private static String getId(String id) {
        String replace = id.replace(".", "");
        return replace.replace("E7", "").trim();
    }

    private static String getFirstName(String firstName) {
        String replace1 = firstName.replace("?", "");
        String replace2 = replace1.replace("\n", "");
        return replace2.replace("❤️", "").trim();
    }

    private static String getLastName(String lastName) {
        String replace1 = lastName.replace("?", "").trim();
        String replace2 = replace1.replace("\n", "");
        return replace2.replace("❤️", "").trim();
    }

    public static String toJSON(List<String> list) {
        StringBuilder sb = new StringBuilder();
        for (String s : list) {
            sb.append(s);
            sb.append("\n");
        }
        return sb.toString();
    }
}
