const express = require('express')
const mongoose = require('mongoose');
const app = express()
const port = 5000

//middleware
app.use(express.json());

var xlsx = require('xlsx');
var workbook = xlsx.readFile('data/employees.xlsx');
let worksheet = workbook.Sheets[workbook.SheetNames[0]];

main().catch(err => console.log(err));

async function main() {
  await mongoose.connect('mongodb://localhost:27017/employee');
  console.log('connection succesfully');
}

const Employee = mongoose.model('Employee', {
    month: { type: String },
    day: { type: Number },
    id: { type: Number },
    employee_name: { type: String },
    department: { type: String },
    first_in_time: { type: Number },
    last_out_time: { type: Number },
    hours_of_works: { type: Number },
});

for(let i = 2; i < 8; i++) {
    const month = worksheet[`A${i}`].v;
    const day = worksheet[`B${i}`].v;
    const id = worksheet[`C${i}`].v;
    const employee_name = worksheet[`D${i}`].v;
    const department = worksheet[`E${i}`].v;
    const first_in_time = worksheet[`F${i}`].v;
    const last_out_time = worksheet[`G${i}`].v;
    const hours_of_works = worksheet[`H${i}`].v;

    const new_employee = new Employee({
        month: month,
        day: day,
        id: id,
        employee_name: employee_name,
        department: department,
        first_in_time: first_in_time,
        last_out_time: last_out_time,
        hours_of_works: hours_of_works
    })

    new_employee.save(function(err,result){
        if (err){
            console.log(err);
        }
        else{
            console.log(result)
        }
    })
}

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
});