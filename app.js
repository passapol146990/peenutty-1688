const express = require('express');
const app = express();
const path = require('path');

app.set("view engine", "ejs");
app.set("views",path.join(__dirname, '/'));

app.use(express.urlencoded({ extended: true, limit: "30mb"}));

app.get('/',(req,res)=>{
    const url = req.query.url;
    res.render('index.ejs',{img: url});
})

app.listen(3000);