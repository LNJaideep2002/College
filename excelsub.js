var mongoose=require('mongoose');
var xl =require("xlsx");
// mongoose.connect('mongodb://localhost/cse',{useNewUrlParser: true, useUnifiedTopology: true});
// var schema_batch=mongoose.Schema({Semester: Number,exam: Array,file: Array});
// var model=mongoose.model('batch',schema_batch);
var noofcos=5;
exam=['CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos1-cos2-]',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#Quize[cos1-]',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#Assignment[cos1-]',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos2-cos3-]',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos3-cos4-]',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos5-]',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#Model[]',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#University[]',
];
var file=['CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos1-cos2-].xlsx',
'CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos2-cos3-].xlsx',
    'CSE#2017-2021#Semester 1#Engineering Mathematics - I#Quize[cos1-].xlsx',
    'CSE#2017-2021#Semester 1#Engineering Mathematics - I#Assignment[cos1-].xlsx',
    'CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos3-cos4-].xlsx',
    'CSE#2017-2021#Semester 1#Engineering Mathematics - I#IAE-Test[cos5-].xlsx',
    'CSE#2017-2021#Semester 1#Engineering Mathematics - I#Model[].xlsx',
    'CSE#2017-2021#Semester 1#Engineering Mathematics - I#University[].xlsx',
    'Pos#CSE#2017-2021#Semester 1#Engineering Mathematics - I.xlsx'
  ];
var data={
    batch: '2017-2021',
    depart: 'CSE',
    semester: 'Semester 1',
    subject: 'Engineering Mathematics - I',
    target: '50'
  }
var target=Number(data.target);
var cos={}
for(i=1;i<=5;i++)
{
    cos["cos"+i]={exams:[],marks:[],div:0,noofstabtar:0,perofatt:0,attlevel:0,coatt:0,poattyorn:"y"};
}
var incos=cos;
if(file[0].search("Pos")==-1)
var wb1=xl.readFile("./public/"+file[0]);
else
var wb1=xl.readFile("./public/"+file[1]);
var ws1=wb1.Sheets["Sheet1"]; 
var data=xl.utils.sheet_to_json(ws1);
for(k in cos)
{
    for(i=0;i<data.length;i++)
    {
        cos[k].marks.push({reg:data[i].Register,mark:0});
    }
}
// console.log(exam[0].split("#")[4].split("["))
for(i=0;i<exam.length;i++)
{
    if(exam[i].search("Model")==-1&&exam[i].search("University")==-1)
    {
      
        var cosin=exam[i].split("#")[4].split("[")[1].split("]")[0].split("-");
        var wb1=xl.readFile("./public/"+exam[i]+".xlsx");
        var ws1=wb1.Sheets["Sheet1"]; 
        var data=xl.utils.sheet_to_json(ws1);
        // console.log(cosin);
        for(t=0;t<cosin.length-1;t++)
        {
            cos[cosin[t]].div++;
            var temp=exam[i].split("#")[4].split("[")[0];
            var f=0;
            for(n=0;n<cos[cosin[t]].exams.length;n++)
            {
                if(temp==cos[cosin[t]].exams[n])
                f=1;
            }
            if(f==0)
            cos[cosin[t]].exams.push(temp)
            for(k=0;k<data.length;k++)
            {
                cos[cosin[t]].marks[k].mark+=data[k].Mark;
            }
        }
    }
    else{
        if(exam[i].search("Model")!=-1)
        {
            var wb1=xl.readFile("./public/"+exam[i]+".xlsx");
            var ws1=wb1.Sheets["Sheet1"]; 
            var data=xl.utils.sheet_to_json(ws1); 
            // console.log(data);
            
            for(j in cos)
            {
                for(k=0;k<cos[j].marks.length;k++)
                {
                    cos[j].marks[k].mark=Math.floor(cos[j].marks[k].mark*1/cos[j].div);
                }
            }
            for(j in cos)
            {
                for(t=0;t<data.length;t++)
                {
                    cos[j].marks[t].mark=Math.floor(cos[j].marks[t].mark*0.8+data[t].Mark*0.2);
                }
            }
        }
       
    }
}
var university={
    noofstabavg:0,                 // No of Students above average
    perofatt:0,
    attlevel:0
} 
var wbg=xl.readFile("./public/Grade.xlsx");
var wsg=wbg.Sheets["grade"]; 
var datag= xl.utils.sheet_to_json(wsg);
var endofsem=[]   // university Marks
var unifilename;
for(i=0;i<file.length;i++)
{
    if(file[i].search("University")!=-1)
    {
        unifilename=file[i]
    }
}
var wbu=xl.readFile("./public/"+unifilename);
var wsu=wbu.Sheets["Sheet1"]; 
var datau= xl.utils.sheet_to_json(wsu);
for(i=0;i<datau.length;i++)
{
    for(j in datag[0])
    {
        if(datau[i].Mark==datag[0][j])
        endofsem.push(Number(j.split('m')[1]));
    }
}
for(i=0;i<endofsem.length;i++)  // calculate the sum of all university marks
    university.noofstabavg+=endofsem[i];
university.noofstabavg/=datau.length;
var uniavg=university.noofstabavg;
university.noofstabavg=0;
for(i=0;i<endofsem.length;i++)
{
    if(endofsem[i]>uniavg)
    university.noofstabavg++;
}
for(i in cos)
{
    for(j=0;j<cos[i].marks.length;j++)
    {
        if(cos[i].marks[j].mark>target)
        cos[i].noofstabtar++;
    }
}
university.perofatt=(university.noofstabavg/(datau.length/100)).toFixed(2);
if(university.perofatt>70)
university.attlevel=3
else if(university.perofatt>60)
university.attlevel=2
else if(university.perofatt>50)
university.attlevel=1
for(i in cos)
{
    cos[i].perofatt=(cos[i].noofstabtar/(datau.length/100)).toFixed(2);
}
for(i in cos)
{
    if(cos[i].perofatt>70)
    cos[i].attlevel=3
   else if(cos[i].perofatt>60)
    cos[i].attlevel=2
    else if(cos[i].perofatt>50)
    cos[i].attlevel=1
}
for(i in cos)
{
    cos[i].coatt=(university.attlevel*0.8+ cos[i].attlevel*0.2).toFixed(2);
    if(cos[i].coatt>=2.5)
    cos[i].poattyorn="Y"
    else
    cos[i].poattyorn="N"
}
var posfile;
for(i=0;i<file.length;i++)
{
    if(file[i].search("Pos")!=-1)
    {
        posfile=file[i]
    }
}
var wbp=xl.readFile("./public/"+posfile);
var wsp=wbp.Sheets["Sheet1"]; 
var datap= xl.utils.sheet_to_json(wsp);
// console.log(datap);
var poattaiment={}; // po attaiment
for(j in datap[0])
{ 
    var posum=0; // to add and multiply
    var po1sum=0; // to add the column po1
    for(i=0;i<datap.length;i++)
    if(j!="COs")
    {
        posum+=cos["cos"+(i+1)].coatt*datap[i][j];
        po1sum+=datap[i][j];
        if(po1sum==0)
        poattaiment[j]=0.00
        else
        poattaiment[j]=(posum/po1sum).toFixed(2);
    }
}
// console.log(poattaiment);
// console.log(university)
console.log(cos)