const express = require('express');
var mongoose = require('mongoose');
const port = 5000;
const app = express();

//middleware
app.use(express.json());

var xlsx = require('xlsx');
var workbook = xlsx.readFile('data/employees.xlsx');

// make a connection 
mongoose.connect('mongodb://localhost:27017/employee');
// get reference to database
var db = mongoose.connection;

db.on('error', console.error.bind(console, 'connection error:'));

let worksheet = workbook.Sheets[workbook.SheetNames[0]];

for(let i = 2; i < 7; i++) {
    const month = worksheet[`A${i}`].v;
    const day = worksheet[`B${i}`].v;
    const id = worksheet[`C${i}`].v;
    const employee_name = worksheet[`D${i}`].v;
    const department = worksheet[`E${i}`].v;
    const first_in_time = worksheet[`F${i}`].v;
    const last_out_time = worksheet[`G${i}`].v;
    const hours_of_works = worksheet[`H${i}`].v;

    console.log({
        month: month,
        day: day,
        id: id,
        employee_name: employee_name,
        department: department,
        first_in_time: first_in_time,
        last_out_time: last_out_time,
        hours_of_works: hours_of_works
    })


    db.once('open', function() {
        console.log("Connection Successful!");
         
        // define Schema
        var EmployeeSchema = mongoose.Schema({
          month: String,
          day: Number,
          id: Number,
          employee_name: String,
          department: String,
          first_in_time: Number,
          last_out_time: Number,
          hours_of_works: Number
        });
     
        // compile schema to model
        var Book = mongoose.model('Book', EmployeeSchema, 'bookstore');
     
        // a document instance
        var book1 = new Book({ name: 'Introduction to Mongoose', price: 10, quantity: 25 });
     
        // save model to database
        book1.save(function (err, book) {
          if (err) return console.error(err);
          console.log(book.name + " saved to employeestore collection.");
        });
         
    });
}

app.listen(port, () => {
    console.log(`Example app listening on port ${port}`)
})