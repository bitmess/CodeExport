package codeexport;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.ui.part.ViewPart;

public class View extends ViewPart {
	public View() {
	}

	public static final String ID = "CodeExport.view";

	@Override
	public void createPartControl(Composite parent) {
		Composite top = new Composite(parent, SWT.NONE);
		GridLayout layout = new GridLayout();
		layout.marginHeight = 0;
		layout.marginWidth = 0;
		top.setLayout(layout);

		Button btnClick = new Button(top, SWT.NONE);
		btnClick.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				try {
					createWord2007();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		});
		btnClick.setText("click");
		

	}

	@Override
	public void setFocus() {
	}

	/**
	 * 2007word文档创建
	 * @throws IOException 
	 */
	public void createWord2007() throws IOException {
		XWPFDocument document= new XWPFDocument();

        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File("/Users/jv/Desktop/create_toc.docx"));

        //添加标题
        XWPFParagraph titleParagraph = document.createParagraph();

        //设置段落居中
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun titleParagraphRun = titleParagraph.createRun();
        titleParagraphRun.setText("Java PoI");
        titleParagraphRun.setColor("000000");
        titleParagraphRun.setFontSize(20);

        //段落
        XWPFParagraph firstParagraph = document.createParagraph();
        firstParagraph.setStyle("Heading1");
        XWPFRun run = firstParagraph.createRun();
        run.setText("段落1。");
        run.setColor("696969");
        run.setFontSize(18);


        //段落
        XWPFParagraph firstParagraph1 = document.createParagraph();
        firstParagraph.setStyle("Heading1");
        XWPFRun run1 = firstParagraph1.createRun();
        run1.setText("段落2");
        run1.setColor("696969");
        run1.setFontSize(16);

        document.createTOC();

        document.write(out);
        out.close();
	}

}
