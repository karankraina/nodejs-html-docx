# nodejs-html-docx

nodejs-html-docx is a simple tool to create MS Word (.Docx) file with Header and Pagination in Footer. That too with Promise. Yayy!
# Installation
```sh
npm i nodejs-html-docx
```

# Usage

```javascript
import {createDoc} from 'nodejs-html-docx'


// bodyMarkup - replaces the content in the body of the document
// headerTitle - Text you want to show in header. Usually a document title
// outputFile - Name you want for the output file

createDoc(bodyMarkup, {headerTitle: 'Report', outputFile: 'Report'}).then(path => {
    // you get the path for the generated file here
    console.log('Conversion complete at => ', path)
}).catch(err => {
    // You know this
	console.log(err)
})
```
