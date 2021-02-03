const express = require('express')
const server = express()
const path = require('path')

const XLSX = require('xlsx')
let workbook = XLSX.readFile('teste.xlsx')
console.log(XLSX)


server.set('view engine', 'ejs')
server.set('views', path.join(__dirname ,'views'))
server.use(express.static( path.join(__dirname, 'public')))


server.get('/', (req, res) => {
    res.render('home')
})

server.listen(3000, err => {
    if(err){
        console.log('Server erro')
    }else{
        console.log('ProcJs Running!')
    }
})