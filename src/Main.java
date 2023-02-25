import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
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
			paragraph.createRun().setText("문제 솰라솰라솰라");

			for (int i = 1; i <= 5; i++) {

				XWPFParagraph seq = document.createParagraph();
				seq.createRun().setText("\t" + i + "번. 설명 솰라솰라.");
			}

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

}