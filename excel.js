var xl =require("xlsx");
var wb1=xl.readFile("iae1.xlsx");
var ws1=wb1.Sheets["Sheet1"]; 
var data1= xl.utils.sheet_to_json(ws1);
var wb2=xl.readFile("iae2.xlsx");
var ws1=wb2.Sheets["Sheet1"]; 
var data2= xl.utils.sheet_to_json(ws1);
var wb3=xl.readFile("iae3.xlsx");
var ws3=wb3.Sheets["Sheet1"]; 
var data3= xl.utils.sheet_to_json(ws3);
var wb4=xl.readFile("iae4.xlsx");
var ws4=wb4.Sheets["Sheet1"]; 
var data4= xl.utils.sheet_to_json(ws4);
var wb5=xl.readFile("iae5.xlsx");
var ws5=wb5.Sheets["Sheet1"]; 
var data5= xl.utils.sheet_to_json(ws5);
var wbm=xl.readFile("model.xlsx");
var wsm=wbm.Sheets["Sheet1"]; 
var datam= xl.utils.sheet_to_json(wsm);
var wbu=xl.readFile("univer.xlsx");
var wsu=wbu.Sheets["Sheet1"]; 
var datau= xl.utils.sheet_to_json(wsu);
var wbg=xl.readFile("Grade.xlsx");
var wsg=wbg.Sheets["grade"]; 
var datag= xl.utils.sheet_to_json(wsg);
var wbp=xl.readFile("pos.xlsx");
var wsp=wbp.Sheets["Sheet1"]; 
var datap= xl.utils.sheet_to_json(wsp);
// console.log(datap);
var endofsem=[];
for(i=0;i<datau.length;i++)
{
    for(j in datag[0])
    {
        if(datau[i].Mark==datag[0][j])
        endofsem.push({endofsem:j.split('m')[1]});
    }
}
var iaf={iae1:[],iae2:[],iae3:[],iae4:[],iae5:[]};
for(i=0;i<data1.length;i++)
{
    iaf.iae1.push({iae1f:Math.floor(data1[i].Mark*0.8+datam[i].Mark*0.2)});
    iaf.iae2.push({iae2f:Math.floor(data2[i].Mark*0.8+datam[i].Mark*0.2)});
    iaf.iae3.push({iae3f:Math.floor(data3[i].Mark*0.8+datam[i].Mark*0.2)});
    iaf.iae4.push({iae4f:Math.floor(data4[i].Mark*0.8+datam[i].Mark*0.2)});
    iaf.iae5.push({iae5f:Math.floor(data5[i].Mark*0.8+datam[i].Mark*0.2)});
}
// console.log(iaf);
var noofstabtar={iae1:0,iae2:0,iae3:0,iae4:0,iae5:0,uni:0}; // no of st above target

for(i=0;i<iaf.iae1.length;i++)
{
if(iaf.iae1[i].iae1f>50) // 50  target
noofstabtar.iae1++;
if(iaf.iae2[i].iae2f>50) // 50  target
noofstabtar.iae2++;
if(iaf.iae3[i].iae3f>50) // 50  target
noofstabtar.iae3++;
if(iaf.iae4[i].iae4f>50) // 50  target
noofstabtar.iae4++;
if(iaf.iae5[i].iae5f>50) // 50  target
noofstabtar.iae5++;
if(endofsem[i].endofsem>50)
noofstabtar.uni++;
}
var perofat={iae1:0,iae2:0,iae3:0,iae4:0,iae5:0,uni:0}; // %of attainment
// console.log(noofstabtar);
perofat.iae1=(noofstabtar.iae1/(data1.length/100)).toFixed(2);
perofat.iae2=(noofstabtar.iae2/(data1.length/100)).toFixed(2);
perofat.iae3=(noofstabtar.iae3/(data1.length/100)).toFixed(2);
perofat.iae4=(noofstabtar.iae4/(data1.length/100)).toFixed(2);
perofat.iae5=(noofstabtar.iae5/(data1.length/100)).toFixed(2);
perofat.uni=(noofstabtar.uni/(data1.length/100)).toFixed(2);
// console.log((43/(60/100)).toFixed(2));
// console.log(perofat);
var atlevel={iae1:0,iae2:0,iae3:0,iae4:0,iae5:0,uni:0};// attainment level
for(i in atlevel)
{

    if(perofat[i]>70)
    atlevel[i]=3
    else if(perofat[i]>60)
    atlevel[i]=2
    else if(perofat[i]>50)
    atlevel[i]=1
}
// console.log(atlevel);
var coat={iae1:0,iae2:0,iae3:0,iae4:0,iae5:0,uni:"-"}; // co attainment
for(i in coat)
{
    coat[i]=(atlevel[i]*0.2+atlevel.uni*0.8).toFixed(2);
}
// console.log(coat);
var attainmentpo={iae1:0,iae2:0,iae3:0,iae4:0,iae5:0}; // attainment of po y or n
for(i in coat)
{
    if(coat[i]>2.5&&i!='uni')
    attainmentpo[i]='y';
    else if (i!='uni')
    attainmentpo[i]='n'
}

var poattaiment={}; // po attaiment
for(j in datap[0])
{ 
    var posum=0; // to add and multiply
    var po1sum=0; // to add the column po1
    for(i=0;i<datap.length;i++)
    if(j!="COs")
    {
        posum+=coat['iae'+(i+1)]*datap[i][j];
        po1sum+=datap[i][j];
        if(po1sum==0)
        poattaiment[j]=0.00
        else
        poattaiment[j]=posum/po1sum;
    }
}
console.log(poattaiment);