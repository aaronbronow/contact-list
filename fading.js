/* $Id: fading.js,v 1.1 2002/10/02 18:49:42 shaggy Exp $ */

/*
Copyright (c) 2001, 2002 by Martin Tsachev. All rights reserved.
mailto:martin@f2o.org
http://martin.f2o.org

Redistribution and use in source and binary forms,
with or without modification, are permitted provided
that the conditions available at
http://www.opensource.org/licenses/bsd-license.html
are met.
*/

var r = g = b = 0;
var toR = toG = toB = 256;
var chR = chG = chB = 2;
var fader = null;
var timer = null;

function setFromColor(R, G, B) {
	r = R;
	g = G;
	b = B;
}

function setToColor(R, G, B) {
	toR = R;
	toG = G;
	toB = B;
}


function fadeIn(obj) {
	if ( obj.style ) { // browser ok
		fader = obj;
		if ( timer )
			clearTimeout(timer);
		fadeReal(chR, chG,chB)
	}
}

function fadeOut(obj) {
	if ( obj.style ) { // browser ok
		fader = obj;
		if ( timer )
			clearTimeout(timer);
		fadeReal(-chR,-chG,-chB)
	}
}

function fade(obj, fR, fG, fB, tR, tG, tB) {
  if( obj.style ) { // browser ok
   fader = obj;
   if( timer )
     clearTimeout(timer);
   r = fR;
   g = fG;
   b = fB;
   toR = tR;
   toG = tG;
   toB = tB;
   fadeReal(chR, chG, chB);
  }
}

function fadeReal(chR, chG, chB) {
	r += chR; // update color values
	g += chG;
	b += chB;

	if ( ( r >= 0 ) && ( r < 256 ) && ( g >= 0 ) && ( g < 256 ) && ( b >= 0 ) && ( b < 256 ) ) {
		fader.style.color = "rgb(" + r + "," + g + "," + b + ")";
		timer = setTimeout("fadeReal(" + chR + "," + chG + "," + chB + ")", 100);
	}
}
