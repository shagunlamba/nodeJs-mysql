require('dotenv').config();
const mysql = require('mysql2');
var q = require('q');
const multer = require('multer');
const express = require("express");
const ejs = require('ejs');
const path = require('path');
const readXlsxFile = require('read-excel-file/node');
const app = express();
const bodyParser = require("body-parser");
app.use(express.static(__dirname + '/public'));

app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'ejs');

var db = mysql.createConnection({    
     host: process.env.DB_HOST,
     user: process.env.DB_USER,
     password: process.env.DB_PASS
     database: "heroku_68b04651207c024"
});

app.get('/', function (req, res) {
    
    var sql='DELETE FROM student';
    db.query(sql, function (err, data, fields) {
    if (err) {
        console.log(err);
    }else{
        res.render('index');
     }
    });
});

// Set storage engine
const storage = multer.diskStorage({
    destination: __dirname + "/public/uploads/",
    filename: function(req,file,callback){
        callback(null,file.fieldname + '-' + Date.now() + path.extname(file.originalname));
    }
});

function checkFileType(file, cb){
    //  checking the file extensions
    //  allowed extensions
    const filetypes = /xlsx|xlsm/;
    const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
    if(extname){
        return cb(null,true);
    }
    else{
        cb("Error: please upload excel sheet only ");
    }
}

// Initialse Upload
const upload = multer({
    storage: storage,
    fileFilter: function(req,file,cb) {
        checkFileType(file,cb);
    }
}).single('myFile');





app.post("/upload", (req,res)=>{

    upload(req,res, (err)=>{
        if(err){
            res.render('index', {msg: err});
        }
        else{
            console.log(req.file);
            // res.send('test');
            if(req.file == undefined){
                res.render('index', {
                    msg: "Error no file selected!"
                });
            }
            else{
                console.log("YAYY GOOOOD");
                var sql='SELECT * FROM student';

                importExcelData2MySQL(__dirname + '/public/uploads/' + req.file.filename).then(function(para){
                    db.query(sql, function (err, data, fields) {
                        if (err) {
                            console.log(err);
                        }
                        else{
                            res.render('answer', { userData: data});
                        }
                    });
                }, function(error){
                    console.log(error);
                }); 
            }
        }
    });
});

// -> Import Excel Data to MySQL database
function importExcelData2MySQL(filePath){

    var sql='DELETE FROM student';
    db.query(sql, function (err, data, fields) {
    if (err) {
        console.log(err);
    }
    });

    var deferred = q.defer();
	readXlsxFile(filePath).then((rows) => {
		// `rows` is an array of rows
		// each row being an array of cells.	 
		console.log(rows);
	
		// Remove Header ROW
		rows.shift();
	 
		// Open the MySQL connection
		db.connect((error) => {
			if (error) {
                console.error(error);
			} else {
				let query = 'INSERT INTO Student (Name, RollNo, Class) VALUES ?';
				db.query(query, [rows], (error, response) => {
                    console.log(error || response);
                    
                    if (error) {
                        //throw err;           
                        deferred.reject(error);
                    }
                    else {
                        //console.log(rows);           
                        deferred.resolve(rows);
                    }
                });
                console.log("yehi pe hu mai");
			}
		});
    })
    return deferred.promise;
}



app.post("/upload/search", function(req,res){
    console.log("Hi");
    console.log(req.body.dataToSearch);
    console.log(req.body.data);
    const dataToSearch = req.body.dataToSearch;
    const dataBySearch = req.body.data;

    var sql= `SELECT * FROM student where ${dataBySearch} like "${dataToSearch}%"`;
    console.log(sql);
    db.query(sql, function (err, data, fields) {
        if (err) {
            console.log(err);
        }
        else{
            console.log(data);
            res.render('search', { userData: data});
        }
    });
});

  setInterval(function () {
    db.query('SELECT 1');
}, 5000);


app.listen(process.env.PORT || 3000, function () {
    console.log("Server is up and running on port 3000");
});
