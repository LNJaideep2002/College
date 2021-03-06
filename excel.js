var xl =require("xlsx");
const generatefile=(iae1,iae2,iae3,iae4,iae5,model,unig,pos,target)=>
{
var wb1=xl.readFile(iae1);
var ws1=wb1.Sheets["Sheet1"]; 
var data1= xl.utils.sheet_to_json(ws1);
var wb2=xl.readFile(iae2);
var ws1=wb2.Sheets["Sheet1"]; 
var data2= xl.utils.sheet_to_json(ws1);
var wb3=xl.readFile(iae3);
var ws3=wb3.Sheets["Sheet1"]; 
var data3= xl.utils.sheet_to_json(ws3);
var wb4=xl.readFile(iae4);
var ws4=wb4.Sheets["Sheet1"]; 
var data4= xl.utils.sheet_to_json(ws4);
var wb5=xl.readFile(iae5);
var ws5=wb5.Sheets["Sheet1"]; 
var data5= xl.utils.sheet_to_json(ws5);
var wbm=xl.readFile(model);
var wsm=wbm.Sheets["Sheet1"]; 
var datam= xl.utils.sheet_to_json(wsm);
var wbu=xl.readFile(unig);
var wsu=wbu.Sheets["Sheet1"]; 
var datau= xl.utils.sheet_to_json(wsu);
var wbg=xl.readFile("./public/Grade.xlsx");
var wsg=wbg.Sheets["grade"]; 
var datag= xl.utils.sheet_to_json(wsg);
var wbp=xl.readFile(pos);
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
// console.log(endofsem);
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
if(iaf.iae1[i].iae1f>target) // target  target
noofstabtar.iae1++;
if(iaf.iae2[i].iae2f>target) // target  target
noofstabtar.iae2++;
if(iaf.iae3[i].iae3f>target) // target  target
noofstabtar.iae3++;
if(iaf.iae4[i].iae4f>target) // target  target
noofstabtar.iae4++;
if(iaf.iae5[i].iae5f>target) // target  target
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
    if(i!="uni")
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

// console.log(attainmentpo);
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
        poattaiment[j]=(posum/po1sum).toFixed(2);
    }
}
// console.log(poattaiment);
var pdfattain=[{name:"No of Students above Target:",...noofstabtar},{name:"% of Attainment",...perofat},{name:"Attainment Level",...atlevel},{name:"CO Attainment(20% of Internal + 80% of External)",...coat}];
var fs=require("fs");
var html1=fs.readFileSync("./html.txt",{encoding:'utf8',flag:'r'});
var html2=fs.readFileSync("./html1.txt",{encoding:'utf8',flag:'r'});
var html3="";
var final=[];
for(i=0;i<data1.length;i++)
final.push({reg:data1[i].Register,iae1:data1[i].Mark,iae2:data2[i].Mark,iae3:data3[i].Mark,iae4:data4[i].Mark,iae5:data5[i].Mark,model:datam[i].Mark,uni:datau[i].Mark,unim:endofsem[i].endofsem,iae1m:iaf.iae1[i].iae1f,iae2m:iaf.iae2[i].iae2f,iae3m:iaf.iae3[i].iae3f,iae4m:iaf.iae4[i].iae4f,iae5m:iaf.iae5[i].iae5f});
for(i=0;i<final.length;i++)
{
   html3+=`<tr height=19 style='height:14.4pt'><td height=19 class=xl724739 style='height:14.4pt;border-top:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${(i+1)}</td>
<td class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;'>${final[i].reg}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae1}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae2}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae3}</td>
<td class=xl754739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae4}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae5}</td>
<td class=xl724739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].model}</td>
<td class=xl764739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:windowtext;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Arial, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;'>${final[i].uni}</td>
<td class=xl724739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].unim}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae1m}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae2m}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae3m}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae4m}</td>
<td class=xl744739 style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;'>${final[i].iae5m}</td>
</tr>`;
}
var html4="";
for(i=0;i<pdfattain.length;i++)
{
    html4+=`    <tr>
    <td colspan="9" class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;padding-left: 100px;'>${pdfattain[i].name}</td>
    <td  class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;'>${pdfattain[i].uni}</td>
    <td  class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;'>${pdfattain[i].iae1}</td>
    <td  class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;'>${pdfattain[i].iae2}</td>
    <td  class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;'>${pdfattain[i].iae3}</td>
    <td  class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;'>${pdfattain[i].iae4}</td>
    <td  class=xl734739 align=center style='border-top:none;border-left:none;padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:10.0pt;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:general;vertical-align:bottom;border:.5pt solid windowtext;'>${pdfattain[i].iae5}</td>
</tr>`
}
var pdf=require("html-pdf");
var html5=fs.readFileSync("./html3.txt",{encoding:'utf8',flag:'r'});
var html6="";
var j=0;
for(i in attainmentpo)
{
    datap[j]["iae"]=attainmentpo[i]
    datap[j]["iaec"]=coat[i];
    j++;
}
for(i=0;i<datap.length;i++)
{html6+=`<tr height=19 style='height:14.4pt'>
  <td height=19 class=xl6511638 style='padding:0px;color:black;font-size:11.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;height:14.4pt;border-top:none'>${datap[i].COs}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].iaec}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].iae}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO1}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO2}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO3}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO4}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO5}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO6}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO7}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO8}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO9}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO10}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO11}</td>
  <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${datap[i].PO12}</td>
 </tr>`
}
html6+="<tr height=19 style='hehright:14.4pt'><td colspan=\"3\" style='border-top:none;;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;font-weight: 700;'>PO Attainment</td>";
var j=0;
for(i in poattaiment)
{
    if(j<12)
    html6+=` <td  style='border-top:none;border-left:none;padding:0px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;'>${poattaiment[i]}</td>`
    j++;
}
html6+="</tr>";
var psoattain=[]; //pso attainment
for(j=0;j<datap.length;j++)
     psoattain.push({pso1:datap[j]["PSO1"],pso2:datap[j]["PSO2"],pso3:datap[j]["PSO3"],iae:datap[j].iae,iaec:datap[j].iaec,cos:datap[j]["COs"]});
     var html7=fs.readFileSync("./html4.txt",{encoding:'utf8',flag:'r'});


     for(i=0;i<psoattain.length;i++)
     {
         html7+=`<tr height=19 style='height:14.4pt'>
         <td height=19 class=xl6513997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;height:14.4pt;border-top:none'>${psoattain[i].cos}</td>
         <td class=xl6713997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${psoattain[i].iaec}</td>
         <td class=xl6713997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${psoattain[i].iae}</td>
         <td class=xl6713997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${psoattain[i].pso1}</td>
         <td class=xl6713997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${psoattain[i].pso2}</td>
         <td class=xl6713997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:400;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${psoattain[i].pso3}</td>
        </tr>
       `   
     }
     html7+=` <tr height=19 style='height:14.4pt'>
     <td colspan=3 height=19 class=xl6813997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:bottom;border:.5pt solid windowtext;white-space:nowrap;height:14.4pt'>PSO Attainment</td>
     <td class=xl6913997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${poattaiment.PSO1}</td>
     <td class=xl6913997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${poattaiment.PSO2}</td>
     <td class=xl6913997 style='padding-top:1px;padding-right:1px;padding-left:1px;color:black;font-size:11.0pt;font-weight:700;font-style:normal;text-decoration:none;font-family:Calibri, sans-serif;text-align:center;vertical-align:middle;border:.5pt solid windowtext;white-space:nowrap;border-top:none;border-left:none'>${poattaiment.PSO3}</td>
    </tr></table>`;
  var html9="<script>function converting(){document.getElementById(\"con\").innerHTML=\"\"; window.print();}</script>"
    var html=html2.split("<body>")[0]+"<body>"+html1+html2.split("<body>")[1]+html3+html4+"</table></div>"+html5+html6+"</table>"+html7+"</body>"+html9+"</html>";
iae1=iae1.split("public/")[1];
var filename="/"+iae1.split("iae1.xlsx")[0]+"Co_and_Po_Attainment.html";
 fs.writeFileSync(filename,html,"utf8");
 return filename;
}
    exports.generatefile=generatefile;
