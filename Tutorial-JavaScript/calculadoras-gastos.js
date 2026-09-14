/*gastos en un año
puedes modificar estos valores para probar diferentes resultados*/
var nombre = "Adrian";
var apellido = "Fernandez";
var banco = 1000;
var ingreso = 1600;
var personas = 3;
//a partir de aquí todo es código
const agua = 160;
const luz = 160;
var par = false;
const casa = 600;
const comida = 200;
const gastosExtra = 400;

for(var m = 1; m <= 12; m++){
	banco += ingreso;
	var gastos = casa + luz + (comida*personas);
	
    if(par){
	    gastos += agua;
	    	par = false;
    } else{
	    par = true;
    }

    if(m == 6 || m == 12){
        banco += ingreso;
        gastos += gastosExtra;
        console.log("Se recibió la paga doble");
    }
    banco -= gastos
    resultado = "${nombre} ${apellido} tuvo ${gastos}€ en gastos en el mes ${m} y tiene ${banco}€ en el banco";
    console.log(resultado);
}
