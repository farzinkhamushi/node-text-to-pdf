const express = require('express');
const { Packer, Document } = require('docx');
const fs = require('fs');
const PDFDocument = require('pdfkit');
const path = require('path');

const app = express();
const PORT = 3000;

// Middleware
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.use(express.static('public'));

// GET route to display the form
app.get('/', (req, res) => {
    res.render('index');
});

// POST route to handle form submission and generate files
app.post('/generate', (req, res) => {
    let inputText = req.body.text; // Get text from form
    let fileName = req.body.name;

    if (!inputText) {
        if(!fileName){
            return res.status(400).send('Please provide file name.');
        }
        return res.status(400).send('Please provide some text.');
    }

    
    //let corect_text = processText(inputText);

    create_word_by_text(fileName , inputText);

    create_pdf_by_text(fileName , inputText);

    send_response_with_download_links(fileName,res);

});


let create_pdf_by_text = (fileName , inputText) => {
    // Generate PDF file
    const pdfFilePath = path.join(__dirname, 'public', `${fileName}.pdf`);
    const docPDF = new PDFDocument({ bufferPages:true });

    // Setfont that support Persian characters
    docPDF.font('fonts/Vazirmatn-Medium.ttf'); // Make sure you have this font in the 'fonts' folder
    
    docPDF.pipe(fs.createWriteStream(pdfFilePath));
    docPDF.fontSize(12).text(inputText, { align: 'center' });
    docPDF.end();
}

let send_response_with_download_links = (fileName,res) => {
    // Send response with download links
    res.render('result', { wordFile: `/${fileName}.docx`, pdfFile: `/${fileName}.pdf`});
}

let StringToArray = (inputString) => {
    // Regular expression to split the string into words and HTML tags
    const regex = /(<[^>]+>|[^<>\s]+)/g;
    return inputString.match(regex);
}

let startsWith = (inputWord , includedChar) => {
    if(includedChar.length == 1){
        let chars = StringToArray(inputWord);
        /*
        let regex = /(<[^>]+>|[^<>\s]+)/g;
        let chars = inputWord.match(regex);
        */
        return (chars[0] == includedChar) ? 'true' : 'false';
    }
}

let reverser = (items) => {
    let reversed_arr = [];
    let j = items.length;
    let i = 0;
    while(items){
        reversed_arr[j] = items[i];
        j--;
        i++;
    }
    return reversed_arr;
}

let processText = (inputString) => {
    let tokens = StringToArray(inputString);
    let final_array = [];  
    let persian_block = [];
    let html_block = [];
    for (let i = 0; i < tokens.length; i++) {
        if(!startsWith(tokens[i],'<') && /[\u0600-\u06FF]/.test(tokens[i])){
            persian_block.push(tokens[i]);
            if(!(i+1)>= tokens.length){
                if( startsWith(tokens[i+1],'<') ){
                    final_array.push(...reverser(persian_block));
                    persian_block = [];
                }
            }else{
                final_array.push(...reverser(persian_block));
                persian_block = [];
            }
        }
        else{
            if( startsWith(tokens[i],'<') ){
                html_block.push(tokens[i]);
            }
            if(!(i+1)>= tokens.length){
                if( !startsWith(tokens[i+1],'<') ){
                    final_array.push(...html_block);
                    html_block = [];
                }
            }else{
                final_array.push(...html_block);
                html_block = [];
            }
        }

    }
    let final_sentence = "" ;
    final_array.forEach(element => {
        final_sentence += element + " ";
    });
    return final_sentence;
}


// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});



let create_word_by_text = (fileName , inputText) => {

    // Generate Word file
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                {
                    text: inputText,
                    paragraph: {
                        alignment: "center",
                        style: "PersianStyle" // Custom style for Persian text
                    },
                },
            ],
        }],
        styles: {
            paragraphStyles: [
                {
                    id: "PersianStyle",
                    name: "Persian Style",
                    run: {
                        font: "B Nazanin", // Use a Persian font
                        size: 24,
                    },
                },
            ],
        },
    });



    const wordFilePath = path.join(__dirname, 'public', `${fileName}.docx`);
    // Use Packer to write the document to a file

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync(wordFilePath, buffer, { encoding: 'utf-8' });
    });

}













/*
    // Helper function to reverse Persian words in a segment
    function reversePersianSegment(segment) {
        // Filter out only Persian words
        let persianWords = segment.filter(token => !token.startsWith('<') && /[\u0600-\u06FF]/.test(token));
        
        // Reverse the order of Persian words
        persianWords.reverse();

        // Replace the original Persian words with reversed ones
        return segment.map(token => 
            token.startsWith('<') ? token : // Keep HTML tags as is
            persianWords.shift() || token   // Replace Persian words or keep non-Persian words
        );
    }
    // Process each token and group them by segments (separated by HTML tags)
    let result = [];
    let currentSegment = [];
    tokens.forEach(token => {
        if (token.startsWith('<')) {
            // If it's an HTML tag, process the current segment and reset
            if (currentSegment.length > 0) {
                result.push(...reversePersianSegment(currentSegment));
                currentSegment = [];
            }
            result.push(token); // Add the HTML tag as is
        } else {
            // Otherwise, add the token to the current segment
            currentSegment.push(token);
        }
    });
    
    // Process the last segment if any
    if (currentSegment.length > 0) {
        result.push(...reversePersianSegment(currentSegment));
    }
    
    let total_sentence = "" ;
    result.forEach(element => {
        total_sentence += element + " ";
    });
    return result.join(' ');
*/




