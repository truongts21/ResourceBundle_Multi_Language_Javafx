package exel;

import java.io.File;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import com.gembox.spreadsheet.CellValueType;
import com.gembox.spreadsheet.ExcelCell;
import com.gembox.spreadsheet.ExcelFile;
import com.gembox.spreadsheet.ExcelWorksheet;
import com.gembox.spreadsheet.SpreadsheetInfo;

public class Demo {

	public static void main(String[] args) {
		SpreadsheetInfo.setLicense("FREE-LIMITED-KEY");
		// File fileName = new File("language.xlsx");
		File fileName = new File("Booking.xlsx");
		// File fileName = new File("OccupancyForecast.xlsx");
		// File fileName = new File("Availability.xlsx");
		// File fileName = new File("RevenueForecast.xlsx");
		// File fileName = new File("DailyPickup.xlsx");

		ExcelFile workbook = null;
		try {
			workbook = ExcelFile.load(fileName.getAbsolutePath());
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		ExcelWorksheet worksheet = workbook.getWorksheets().getActiveWorksheet();

		int rowsSize = worksheet.getRows().size();
		String field = "";
		String nameEn = "";
		String nameVi = "";
		StringBuffer buffConstants = new StringBuffer();
		StringBuffer buffEn = new StringBuffer();
		StringBuffer buffVi = new StringBuffer();

		Path fileVn = Paths.get("src/bundle/Languge_vi.properties");
		List<String> content = new ArrayList<String>();
		try {
			content = Files.readAllLines(fileVn, StandardCharsets.UTF_8);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		for (String str : content) {
			buffVi.append(str).append("\n");
		}

		for (int row = 1; row < rowsSize; row++) {

			ExcelCell cell = worksheet.getCell(row, 1);
			if (cell.getValueType() != CellValueType.NULL) {
				field = cell.getValue().toString();
			}
			cell = worksheet.getCell(row, 2);
			if (cell.getValueType() != CellValueType.NULL) {
				nameEn = cell.getFormattedValue();
			}
			cell = worksheet.getCell(row, 3);
			if (cell.getValueType() != CellValueType.NULL) {
				nameVi = cell.getValue().toString();
			}
			String constName = fieldToConstant(field);
			if (!foundConstant(constName)) {
				buffConstants.append("public static final String " + constName + " = \"" + field + "\";\n");
				buffEn.append(field + " = " + nameEn + "\n");
				buffVi.append(field + " = " + nameVi + "\n");
			}
		}
		System.out.println(buffConstants);
		System.out.println();
		System.out.println(buffEn);
		System.out.println();
		System.out.println(buffVi);

		try {
			Files.write(fileVn, buffVi.toString().getBytes("UTF8"));

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static String fieldToConstant(String input) {
		char[] chars = input.toCharArray();
		String pre = null;
		StringBuffer buffer = new StringBuffer();
		for (char ch : chars) {
			String curr = String.valueOf(ch);

			if (pre == null) {
				buffer.append("TL_");
			} else {
				if (pre.equals(pre.toLowerCase()) && !curr.equals(curr.toLowerCase())) {
					buffer.append("_");
				}
			}
			buffer.append(curr.toUpperCase());

			pre = curr;
		}

		return buffer.toString();
	}

	private static boolean foundConstant(String name) {
		Field[] fields = Translate.class.getDeclaredFields();

		for (Field field : fields) {
			if (field.getName().equals(name)) {
				return true;
			}
		}

		return false;
	}

}
