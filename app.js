const express = require("express");
const bodyParser = require("body-parser");
const mongoose = require("mongoose");
const Excel = require('exceljs');
const path = require('path');
const fs = require('fs');
var cors = require('cors');


const app = express();
app.use(bodyParser.urlencoded({extended:false}));

app.use(cors());
app.use(express.json())

mongoose.connect("mongodb+srv://keshavsomani999:POfc3CKGuOMO5CVG@cluster0.g5pdodt.mongodb.net/?retryWrites=true&w=majority",{useNewUrlParser:true,useUnifiedTopology:true}).then(()=>{
    console.log(`Connected with mongodb`);
}).catch(()=>{
    console.log("Error");
})

const current = new Date().getUTCDate();

const studentSchema = mongoose.Schema({
    name:String,
    reg:String,
    block:String,
    createdAt:{type:Date,default:Date.now()}
})

const Student = new mongoose.model("Student",studentSchema)


app.post("/api/v1",async(req,res)=>{

    const student = await Student.create(req.body);

    res.status(200).json({
        success:true,
        student,
    })
})




  



const exportCountriesFile = async (data) => {
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Students List');
  
    worksheet.columns = [
      { key: 'name', header: 'Name' },
      { key: 'reg', header: 'Registration No.' },
      { key: 'block', header: 'Block' },
    ];
  
    data.forEach((item) => {
      worksheet.addRow(item);
    });
  
    const p = './Students.xlsx';
    
    // let exportPath ;
        //    const exportPath = path.resolve(__dirname, 'Students.xlsx');

        // var wopts = {bookType: 'xlsx', bookSST: true, type: 'binary', cellStyles: true};
   
        if(fs.existsSync(p)){
            // fs.unlinkSync(p);
            //  exportPath = path.resolve(__dirname, 'Students.xlsx');
            await workbook.xlsx.writeFile(p);
        }
        else{
            exportPath = path.resolve(__dirname, 'Students.xlsx');
            await workbook.xlsx.writeFile(exportPath);

        }
    
  };


  app.get("/",async(req,res)=>{

    res.status(200).send("buy")

  })


app.get("/api/v1/admin",async(req,res)=>{
  window.print("01")
    res.setHeader('Content-Type', 'application/vnd.openxmlformats');
    res.setHeader("Content-Disposition", "attachment; filename=" + 
    "Students.xlsx");
    console.log("02");
    let date = new Date().toISOString();
    let result = date.slice(0,10);
    console.log(result);

    const data = await Student.find({createdAt:{$gte:`${result}T00:01:00.597+00:00`}})
    res.status(500).send('IServerr');

    await exportCountriesFile(data);
    const filePath = path.join(__dirname, 'Students.xlsx');

    // res.download(filePath, (err) => {
    //     if (err) {
    //       console.error(err);
    //       res.status(500).send('Internal Server Error');
    //     } else {
    //       console.log('File sent successfully!');
          
    //     }
    //   });

    
})
// ll

app.listen(4500,() => {
    console.log(`Server is working http://localhost:4500`);
})