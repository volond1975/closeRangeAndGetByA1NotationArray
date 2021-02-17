String.prototype.to10 = function(base) {
  var lvl = this.length - 1;
  var val = (base || 0) + Math.pow(26, lvl) * (this[0].toUpperCase().charCodeAt() - 64 - (lvl ? 0 : 1));
  return (this.length > 1) ? (this.substr(1, this.length - 1)).to10(val) : val;
}


Number.prototype.to26 = function(suffix) {
    suffix = String.fromCharCode((this % 26) + 65) + (suffix || '');
    return this >= 26 ? (Math.floor(this / 26) - 1).to26(suffix) : suffix;
};

/**
    * Для парсинга инфы получаем или весь документ
    * или ячейку на основании формурмулы
    * @param {*} values
    * @param {*} formulaText
    * @returns
    */
function getValueByFormulaText(values, formulaText) {
    if (formulaText) {
        var [a, ...rest] = [...formulaText]
        var c = +rest.filter(l => /\D/.test(l)).join("").to10()
        var r = +rest.filter(l => /\d/.test(l)).join("")
        var value = values[r - 1][c]
    } else {
        var value = values.join("\n")

    }

    return value
}

var isColon=(A1Not)=>!!([...A1Not].indexOf(":")+1)
function getA1Nots(A1Not){
var [start,end]=isColon(A1Not)?A1Not.split(":"):[A1Not,A1Not]
return [start,end]
}

function parce(formulaText){
let rest=[...formulaText]
const filter=(rest,reg)=>rest.filter(l => reg.test(l))
let filterRestR=filter(rest,/\D/)
let c=filterRestR.length?(+filterRestR.join("").to10()):undefined
let filterRestC=filter(rest,/\d/)
let r=filterRestC.length?(+filterRestC.join("")):undefined
let res={c,r}
      for (val of Object.keys(res)) {
          if(res[val]===undefined)  delete res[val]
          }
  
        return res

}



function closeA1Not(A1NotDataRange,A1Not){
let dataRangeObj=getA1Nots(A1NotDataRange).map(parce)
let rangeObj=getA1Nots(A1Not).map(parce).map((el,i)=>{return {...dataRangeObj[i],...el}})
let A1NotClosed=rangeObj.map(el=>el.c.to26()+el.r).join(":")
return {rangeObj,A1NotClosed}
}

var testCloseRange=()=>{
var A1NotDataRange="A1:E7"
var A1NotRow="2:3"
var A1NotColumn="B:C"
var A1NotOne="F10"
var A1NotRange1="B2:C"
var A1NotRange2="B2:3"
var {parse:jp,stringify:js}=JSON
//console.log('A1NotRow '+ js(closeA1Not(A1NotDataRange,A1NotRow)))
//console.log('A1NotColumn '+js( closeA1Not(A1NotDataRange,A1NotColumn)))
//console.log('A1NotOne '+ js(closeA1Not(A1NotDataRange,A1NotOne)))
//console.log('A1NotRange1 '+ js(closeA1Not(A1NotDataRange,A1NotRange1)))
console.log('A1NotRange2 '+ js(closeA1Not(A1NotDataRange,A1NotRange2)))
/** Получаем закрытые диапазоны на основании sheet.getDataRange.getA1Notation()
 *  Использование
 *  Пользователь выделил несколько строк 
 *  Мы же используем ТОЛЬКО ячейки которые выделены
 *  в DATARANGE
 */
/*
Результат тестовых данных
A1NotRow {"rangeObj":[{"c":0,"r":2},{"c":4,"r":3}],"A1NotClosed":"A2:E3"}
A1NotColumn {"rangeObj":[{"c":1,"r":1},{"c":2,"r":7}],"A1NotClosed":"B1:C7"}
A1NotOne {"rangeObj":[{"c":5,"r":10},{"c":5,"r":10}],"A1NotClosed":"F10:F10"}
A1NotRange1 {"rangeObj":[{"c":1,"r":2},{"c":2,"r":7}],"A1NotClosed":"B2:C7"}
A1NotRange2 {"rangeObj":[{"c":1,"r":2},{"c":4,"r":3}],"A1NotClosed":"B2:E3"}
*/
/** Получаем дапазон в Нотации А1 из массива values
 *  Использование
 *  Первым запросом получаем Все данный листа из DATARANGE
 *  потом нам надо получать выборку из ячеек этого диапвзона
 *  Поэтому мы закешируем первый результат в масив
 *  и будем получать нужные данные но используя натацию А1
 *  используя функцию closeA1Not
*/

var values=[[11,12,13,14,15],[21,22,23,24,25],[31,32,33,34,35],[41,42,43,44,45],[51,5,5,5,5],[6,6,6,6,6],[7,7,7,7,7]]  
var filterBuRow=({rangeObj},row,i)=>(row,i)=>{
return (i>=rangeObj[0]['r'])&&(i<=rangeObj[1]['r'])
}
var filterBuCol=({rangeObj},cell,i)=>(cell,i)=>{
  return (i>=rangeObj[0]['c'])&&(i<=rangeObj[1]['c'])
  }
var closedObj=closeA1Not(A1NotDataRange,A1NotRange2)
var fr=filterBuRow(closedObj)
var fc=filterBuCol(closedObj)
console.log(values.filter(fr).map(row=>row.filter(fc)))
/**
 * Результат
 * Інформація	[ [ 32, 33, 34, 35 ], [ 42, 43, 44, 45 ] ]
 */
}
