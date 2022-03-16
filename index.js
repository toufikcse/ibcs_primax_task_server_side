const express = require('express')
const mongoose = require('mongoose');
const cors = require('cors');
const app = express()
const port = 5000

//middleware
app.use(express.json());
app.use(cors());

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
    date: { type: String },
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
    const date = worksheet[`C${i}`].v;
    const id = worksheet[`D${i}`].v;
    const employee_name = worksheet[`E${i}`].v;
    const department = worksheet[`F${i}`].v;
    const first_in_time = worksheet[`G${i}`].v;
    const last_out_time = worksheet[`H${i}`].v;
    const hours_of_works = worksheet[`I${i}`].v;

    const new_employee = new Employee({
        month: month,
        day: day,
        date: date,
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

app.get("/", async (req, res) => {
    const result = await Employee.find({});
    res.send(result);
})

app.get("/:id", async (req, res) => {
    const id = req.params.id;
    console.log("hitting find id", id);
    const result = await Employee.find({id: `${id}`}).exec();
    res.send(result);
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
});