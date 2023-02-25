import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Main {

	public static void main(String[] args) throws FileNotFoundException, Exception {
		try {
			// Word 문서 생성
			XWPFDocument document = new XWPFDocument();

			// 문단 생성
			XWPFParagraph paragraph = document.createParagraph();
			extracted(document).setText("(객관식)");

			for (int i = 1; i <= 5; i++) {
				extracted(document).setText(i + ".ooooooooooooooooooooooooooooooooooooooo 문제 설명");
				extracted(document).setText("-----------------------------------------------");
				extracted(document).setText("1.ooooooooooooooooo\t3.ooooooooooooooooo");
				extracted(document).setText("1.ooooooooooooooooo\t3.ooooooooooooooooo");
				extracted(document).setText("2.ooooooooooooooooo\t4.ooooooooooooooooo");
				extracted(document).setText("5.ooooooooooooooooo");
			}

			extracted(document).setText("\n\n\n(서술형)");
			extracted(document).setText("1.문제 설명 : ooooooooooooooooooooooooooooooooooooooooooo");
			extracted(document).setText("ooooooooooooooooooooooooooooooooooooooooooo");
			extracted(document).setText("서술 : ooooooooooooooooooooooooooooooooooooooooooo");
			extracted(document).setText("ooooooooooooooooooooooooooooooooooooooooooo");
			// 표 생성
			XWPFTable table = document.createTable();

			XWPFTableRow row = table.getRow(0);
			row.getCell(0).setText("첫 번째 셀");
			row.addNewTableCell().setText("두 번째 셀");
			table.setTopBorder(XWPFBorderType.SINGLE, 0, 0, "");

			// 파일 저장
			FileOutputStream fos = new FileOutputStream("example.docx");
			document.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static XWPFRun extracted(XWPFDocument document) {
		return document.createParagraph().createRun();
	}

}