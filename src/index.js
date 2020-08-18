var HtmlDocx = require('html-docx-js');
var fs = require('fs');

const generateMarkup = (markupText, { headerTitle }) => {
    return `<html>
  <head>
  <style>
  table {
	border-collapse: collapse;
  }  
  table, td, th {
	border: 1px solid black;
  }
  </style>
	  <style type="text/css">
	  @page Section1 {
		  margin:0.75in 0.75in 0.75in 0.75in;
		  size:841.7pt 595.45pt;
		  mso-page-orientation:landscape;
		  mso-header-margin:0.5in;
		  mso-header: h1;
		  mso-footer-margin:0.5in;
		  mso-footer: f1;
	  }
  
	  div.Section1 {page:Section1;}
  
	  p.headerFooter { margin:0in; text-align: center; }
	  </style>
  </head>
  <body><div class=Section1>
  
  
  <!-- header/footer:
	This element will appears in your main document (unless you save in a separate HTML),
	therefore, we move it off the page (left 50 inches) and relegate its height
	to 1pt by using a table with 1 exact-height row
  -->
  <table style='margin-left:50in;'><tr style='height:1pt;mso-height-rule:exactly'>
	  <td>
		<div style='mso-element:header' id=h1>
		  <p class=headerFooter>
			   <b>${headerTitle}</b>
		   </p>
		</div>
		&nbsp;
	  </td>
  
	  <td>
		<div style='mso-element:footer' id=f1>
		  <p class=headerFooter>
		  Page
		  <span style='mso-field-code:PAGE'></span>
		  of
		  <span style='mso-field-code:NUMPAGES'></span>
		</p>
		</div>
		&nbsp;
  </td></tr></table>
  
  ${markupText}
  <!-- Here's a page break:
  <br clear=all style='mso-special-character:line-break; page-break-before:always'>
  This is page 2 -->
  
  </div></body>
  </html>`
}

export const createDoc = async (markupText, { headerTitle, outputFile }) => {
    return new Promise((resolve, reject) => {
        const inputMarkup = generateMarkup(markupText, { headerTitle })
        var outputFileName = `${outputFile || 'Report'}.docx`;
        console.log({ outputFileName })
        var docx = HtmlDocx.asBlob(inputMarkup);
        console.log({ docx })
        fs.writeFile(outputFileName, docx, function (err) {
            if (err) {
                reject(err)
                return 0
            };
            resolve(`${outputFileName}`)
        });
    })

}