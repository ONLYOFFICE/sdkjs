/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */
 
function t(){var a={root:{root:!0,next:null},S:function(e){return null===e||e===a.root?!1:!0},ea:function(){return null===a.root.next},N:function(){return a.root.next},insertBefore:function(e,c){for(var d=a.root,b=a.root.next;null!==b;){if(c(b)){e.m=b.m;e.next=b;b.m.next=e;b.m=e;return}d=b;b=b.next}d.next=e;e.m=d;e.next=null},ba:function(e){for(var c=a.root,d=a.root.next;null!==d&&!e(d);)c=d,d=d.next;return{before:c===a.root?null:c,after:d,ca:function(b){b.m=c;b.next=d;c.next=b;null!==d&&(d.m=b);
return b}}}};return a}function F(a){a.m=null;a.next=null;a.remove=function(){if(a.m.next=a.next)a.next.m=a.m;a.m=null;a.next=null};return a}
function G(a){function e(f,l,h){return{id:-1,start:f,end:l,h:{i:h.h.i,j:h.h.j},o:null}}function c(f,l){m.insertBefore(f,function(h){var r=f.L,g=h.L,A=h.v;h=h.A.v;var k=n.V(f.v,A);return 0>(0!==k?k:n.s(l,h)?0:r!==g?r?1:-1:n.P(l,g?A:h,g?h:A)?1:-1)})}function d(f,l){var h=F({L:!0,v:f.start,g:f,D:l,A:null,status:null});c(h,f.end);f=F({L:!1,v:f.end,g:f,D:l,A:h,status:null});h.A=f;c(f,h.v);return h}function b(f,l){var h=e(l,f.g.end,f.g);f.A.remove();f.g.end=l;f.A.v=l;c(f.A,f.v);return d(h,f.D)}function p(f,
l){function h(v){return g.ba(function(u){var q=u.I;u=v.g.start;var w=v.g.end,z=q.g.start;q=q.g.end;return 0<(n.B(u,z,q)?n.B(w,z,q)?1:n.P(w,z,q)?1:-1:n.P(u,z,q)?1:-1)})}function r(v,u){var q=v.g,w=u.g,z=q.start;q=q.end;var E=w.start;w=w.end;var y=n.fa(z,q,E,w);if(!1===y){if(!n.B(z,q,E)||n.s(z,w)||n.s(q,E))return!1;y=n.s(z,E);var J=n.s(q,w);if(y&&J)return u;var N=!y&&n.U(z,E,w);E=!J&&n.U(q,E,w);if(y)return E?b(u,q):b(v,w),u;N&&(J||(E?b(u,q):b(v,w)),b(u,z))}else 0===y.F&&(-1===y.G?b(v,E):0===y.G?b(v,
y.v):1===y.G&&b(v,w)),0===y.G&&(-1===y.F?b(u,z):0===y.F?b(u,y.v):1===y.F&&b(u,q));return!1}for(var g=t(),A=[];!m.ea();){var k=m.N();if(k.L){var x=h(k),H=x.before?x.before.I:null,B=x.after?x.after.I:null,C=function(){if(H){var v=r(k,H);if(v)return v}return B?r(k,B):!1}();if(C){if(a){var D;if(D=null===k.g.h.j?!0:k.g.h.i!==k.g.h.j)C.g.h.i=!C.g.h.i}else C.g.o=k.g.h;k.A.remove();k.remove()}if(m.N()!==k)continue;a?(D=null===k.g.h.j?!0:k.g.h.i!==k.g.h.j,k.g.h.j=B?B.g.h.i:f,D?k.g.h.i=!k.g.h.j:k.g.h.i=k.g.h.j):
null===k.g.o&&(C=B?k.D===B.D?B.g.o.i:B.g.h.i:k.D?l:f,k.g.o={i:C,j:C});k.A.status=x.ca(F({I:k}))}else x=k.status,null===x&&console.log("error"),g.S(x.m)&&g.S(x.next)&&r(x.m.I,x.next.I),x.remove(),k.D||(x=k.g.h,k.g.h=k.g.o,k.g.o=x),A.push(k.g);m.N().remove()}return A}var n=I,m=t();return a?{Y:function(f){for(var l,h=f[f.length-1],r=0;r<f.length;r++){l=h;h=f[r];var g=n.V(l,h);0!==g&&d({id:-1,start:0>g?l:h,end:0>g?h:l,h:{i:null,j:null},o:null},!0)}},M:function(f){return p(f,!1)}}:{M:function(f,l,h,r){f.forEach(function(g){d(e(g.start,
g.end,g),!0)});h.forEach(function(g){d(e(g.start,g.end,g),!1)});return p(l,r)}}}
function K(a){var e=I,c=[],d=[];a.forEach(function(b){function p(B,C,D){r.index=B;r.C=C;r.O=D;if(r===l)return r=h,!1;r=null;return!0}function n(B,C){var D=c[B],v=c[C],u=D[D.length-1],q=D[D.length-2],w=v[0],z=v[1];e.B(q,u,w)&&(D.pop(),u=q);e.B(u,w,z)&&v.shift();c[B]=D.concat(v);c.splice(C,1)}var m=b.start,f=b.end;if(e.s(m,f))console.warn("PolyBool: Warning: Zero-length segment detected; your epsilon is probably too small or too large");else{for(var l={index:0,C:!1,O:!1},h={index:0,C:!1,O:!1},r=l,g=
0;g<c.length;g++){b=c[g];var A=b[0];b=b[b.length-1];if(e.s(A,m)){if(p(g,!0,!0))break}else if(e.s(A,f)){if(p(g,!0,!1))break}else if(e.s(b,m)){if(p(g,!1,!0))break}else if(e.s(b,f)&&p(g,!1,!1))break}if(r===l)c.push([m,f]);else if(r===h){g=l.index;m=l.O?f:m;f=l.C;b=c[g];A=f?b[0]:b[b.length-1];var k=f?b[1]:b[b.length-2],x=f?b[b.length-1]:b[0],H=f?b[b.length-2]:b[1];e.B(k,A,m)&&(f?b.shift():b.pop(),A=k);e.s(x,m)?(c.splice(g,1),e.B(H,x,A)&&(f?b.pop():b.shift()),d.push(b)):f?b.unshift(m):b.push(m)}else b=
l.index,g=h.index,m=c[b].length<c[g].length,l.C?h.C?m?(c[b].reverse(),n(b,g)):(c[g].reverse(),n(g,b)):n(g,b):h.C?n(b,g):m?(c[b].reverse(),n(g,b)):(c[g].reverse(),n(b,g))}});return d}function L(a,e){var c=[];a.forEach(function(d){var b=(d.h.i?8:0)+(d.h.j?4:0)+(d.o&&d.o.i?2:0)+(d.o&&d.o.j?1:0);0!==e[b]&&c.push({id:-1,start:d.start,end:d.end,h:{i:1===e[b],j:2===e[b]},o:null})});return c}
var I=function(a){"number"!==typeof a&&(a=1E-10);var e={R:function(c){"number"===typeof c&&(a=c);return a},P:function(c,d,b){var p=d[0];d=d[1];return(b[0]-p)*(c[1]-d)-(b[1]-d)*(c[0]-p)>=-a},U:function(c,d,b){var p=b[0]-d[0];b=b[1]-d[1];c=(c[0]-d[0])*p+(c[1]-d[1])*b;return c<a||c-(p*p+b*b)>-a?!1:!0},W:function(c,d){return Math.abs(c[0]-d[0])<a},X:function(c,d){return Math.abs(c[1]-d[1])<a},s:function(c,d){return e.W(c,d)&&e.X(c,d)},V:function(c,d){return e.W(c,d)?e.X(c,d)?0:c[1]<d[1]?-1:1:c[0]<d[0]?
-1:1},B:function(c,d,b){return Math.abs((c[0]-d[0])*(d[1]-b[1])-(d[0]-b[0])*(c[1]-d[1]))<a},fa:function(c,d,b,p){var n=d[0]-c[0];d=d[1]-c[1];var m=p[0]-b[0],f=p[1]-b[1];p=n*f-d*m;if(Math.abs(p)<a)return!1;var l=c[0]-b[0];b=c[1]-b[1];m=(m*b-f*l)/p;b=(n*b-d*l)/p;c={F:0,G:0,v:[c[0]+m*n,c[1]+m*d]};c.F=m<=-a?-2:m<a?-1:m-1<=-a?0:m-1<a?1:2;c.G=b<=-a?-2:b<a?-1:b-1<=-a?0:b-1<a?1:2;return c},pa:function(c,d){var b=c[0];c=c[1];for(var p=d[d.length-1][0],n=d[d.length-1][1],m=!1,f=0;f<d.length;f++){var l=d[f][0],
h=d[f][1];h-c>a!=n-c>a&&(p-l)*(c-h)/(n-h)+l-b>a&&(m=!m);p=l;n=h}return m}};return e}(void 0);window.AscGeometry=window.T=window.AscGeometry||{};
var O={R:function(a){return I.R(a)},u:function(a){var e=G(!0);a.regions.forEach(e.Y);return{u:e.M(a.l),l:a.inverted}},Z:function(a,e){return{H:G(!1).M(a.u,a.l,e.u,e.l),J:a.l,K:e.l}},la:function(a){return{u:L(a.H,[0,2,1,0,2,2,0,0,1,0,1,0,0,0,0,0]),l:a.J||a.K}},ka:function(a){return{u:L(a.H,[0,0,0,0,0,2,0,2,0,0,1,1,0,2,1,0]),l:a.J&&a.K}},ia:function(a){return{u:L(a.H,[0,0,0,0,2,0,2,0,1,1,0,0,0,1,2,0]),l:a.J&&!a.K}},ja:function(a){return{u:L(a.H,[0,2,1,0,0,0,1,1,0,2,0,2,0,0,0,0]),l:!a.J&&a.K}},ma:function(a){return{u:L(a.H,
[0,2,1,0,2,0,0,1,1,0,0,2,0,1,2,0]),l:a.J!==a.K}},ga:function(a){a={ha:K(a.u),l:a.l};a.regions=a.ha;a.inverted=a.l;return a},na:function(a,e){return M(a,e,O.la)},da:function(a,e){return M(a,e,O.ka)},$:function(a,e){return M(a,e,O.ia)},aa:function(a,e){return M(a,e,O.ja)},xor:function(a,e){return M(a,e,O.ma)}};O.union=O.na;O.intersect=O.da;O.difference=O.$;O.differenceRev=O.aa;O.xor=O.xor;function M(a,e,c){a=O.u(a);e=O.u(e);e=O.Z(a,e);c=c(e);return O.ga(c)}window.T.PolyBool=window.T.oa=O;
