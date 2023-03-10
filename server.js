// imports
const https = require('https');
const express = require('express');
const cloudinary = require('cloudinary').v2;
const multer = require('multer');
const path = require("path");
const xlsx = require('xlsx');
const fs = require('fs');
require('dotenv').config();
const { v4: uuidv4 } = require('uuid');
const {stringToUTC} = require("./utils/utils");

////////////////////////////// server ///////////////////////////////
const app = express();
const PORT = process.env.PORT || 8000;
app.listen(PORT);

/////////////////////// Middleware //////////////////////////////////////
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

///////////////////////////////////////////// Configurations /////////////////////////////
// Cloudinary
cloudinary.config({
    cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
    api_key: process.env.CLOUDINARY_API_KEY,
    api_secret: process.env.CLOUDINARY_API_SECRET
});

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/')
    },
    filename: function (req, file, cb) {
        const name = file.fieldname;
        const extension = path.extname(file.originalname);
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9)
        cb(null, `${Date.now()}-${name}-${uniqueSuffix}${extension}`)
    }
})
const upload = multer({ storage: storage });

/////////////////////////////// Routes //////////////////////////////////////////////
//Home Route
app.get('/', (req, res)=>{
    res.status(200).json({message: 'Welcome to Convert Sheet'})
})


//File Upload Route
app.post('/upload', upload.single('file'), async (req, res)=>{
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9)
    const _publicID = `${Date.now()}-${req.file.fieldname}-${uniqueSuffix}`;
    const filePath = req.file.path;


    const uploadRes = cloudinary.uploader.upload(filePath, {
        public_id: _publicID,
        resource_type: 'raw',
        format: 'xlsx'
    });

    uploadRes.then((data) => {
        res.status(200).json({message: 'Data uploaded successfully', data: data.secure_url});
    }).catch((err) => {
        res.status(500).json({message: 'Error uploading data', err});
    });
})

app.get('/convert/:secureUrl', (req, res)=>{

    // Fetch from the secure URl
    https.get('https://res.cloudinary.com/dianptniw/raw/upload/v1677197823/1677197821712-file-1677197821712-154357189.xlsx',
        function(chunkRes) {
        const chunks = [];

        chunkRes.on('data', function(chunk) {
                chunks.push(chunk);
            });

        chunkRes.on('end', function() {
            const buffer = Buffer.concat(chunks);
            const workbook = xlsx.read(buffer, { type: 'buffer' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = xlsx.utils.sheet_to_json(worksheet, { header: 2});

            //Get the birthdays & the names of the participants
            const finalChunk = data.map((profile, idx)=>{
                const fullName = `${profile['First Name']} ${profile['Last Name (Surname)']}`;
                const birthDay = `${profile['Birthday']}`;
                return{fullName, birthDay}
            })

            // Create ICS file
const iCSEvents = `BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//Example Corp//EN
CALSCALE:GREGORIAN
METHOD:PUBLISH
${finalChunk.map((chunk, idx) => {
                const dateInitializer = new Date();
                const fullYear = dateInitializer.getFullYear()
                const currentMonth = dateInitializer.getMonth();
                const currentDay = dateInitializer.getDate();
                console.log(currentMonth, currentDay)
                let dateArr = chunk.birthDay.split('/');
                let month = +dateArr[0];
                let day = +dateArr[1];
                if (idx > 3) {
                    return '';
                }
                return `
BEGIN:VEVENT
UID:${uuidv4(undefined, undefined, undefined)}
DTSTAMP:${stringToUTC(fullYear, currentMonth + 1, currentDay)}
DTSTART:${stringToUTC(fullYear, month, day)}
DURATION:PT5M
SUMMARY:${chunk.fullName}
DESCRIPTION:Today is ${chunk.fullName}'s birthday.
BEGIN:VALARM
TRIGGER:-PT5M
ACTION:DISPLAY
DESCRIPTION:Reminder
END:VALARM
END:VEVENT`;
}).join('')}
END:VCALENDAR`

            const finalFileName = `${Date.now()}-${Math.round(Math.random() * 1E9)}`;

            fs.writeFile(`conversions/${finalFileName}.ics`, iCSEvents, function (err) {
                if (err) throw err;
                console.log('Saved!');
            });

            res.status(200).json({message: 'Conversion Complete', data: iCSEvents})
        });
    }).on('error', function(err) {
        res.status(500).json({message: 'Conversion Unsuccessful', error: err});
    });


})