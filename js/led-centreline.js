(function(root, factory) {
  var api = factory();
  // CEP exposes CommonJS globals inside a browser page. Always publish the
  // browser global there; otherwise the panel silently sees no engine.
  if(root && root.document) root.LedCentreline = api;
  if(typeof module === 'object' && module.exports) module.exports = api;
  else if(root) root.LedCentreline = api;
}(this, function() {
  'use strict';
  var SQRT2 = Math.sqrt(2);
  function finite(n) {return typeof n === 'number' && isFinite(n);}
  function pointInContours(x, y, contours) {
    var inside = false;
    for(var c=0;c<contours.length;c++) {
      var p=contours[c].points || contours[c];
      for(var i=0,j=p.length-1;i<p.length;j=i++) {
        var a=p[i], b=p[j];
        if(((a.y>y)!==(b.y>y)) && x < (b.x-a.x)*(y-a.y)/(b.y-a.y)+a.x) inside=!inside;
      }
    }
    return inside;
  }
  function boundsOf(contours) {
    var b={left:Infinity,top:-Infinity,right:-Infinity,bottom:Infinity};
    for(var c=0;c<contours.length;c++) for(var i=0,p=contours[c].points||contours[c];i<p.length;i++) {
      var q=p[i]; if(!finite(q.x)||!finite(q.y)) continue;
      if(q.x<b.left)b.left=q.x;if(q.x>b.right)b.right=q.x;if(q.y>b.top)b.top=q.y;if(q.y<b.bottom)b.bottom=q.y;
    }
    return finite(b.left)?b:null;
  }
  function rasterize(contours, options) {
    options=options||{}; var b=boundsOf(contours); if(!b) throw new Error('No finite contours.');
    var desired=Math.max(0.1,Number(options.cellMm)||1), budget=Math.max(1000,Number(options.maxCells)||250000);
    var w=Math.max(desired,b.right-b.left), h=Math.max(desired,b.top-b.bottom);
    var cell=Math.max(desired,Math.sqrt((w*h)/budget));
    var cols=Math.max(3,Math.ceil(w/cell)+2), rows=Math.max(3,Math.ceil(h/cell)+2);
    var originX=b.left-cell, originY=b.bottom-cell, mask=new Uint8Array(cols*rows);
    for(var y=0;y<rows;y++) for(var x=0;x<cols;x++) mask[y*cols+x]=pointInContours(originX+(x+.5)*cell,originY+(y+.5)*cell,contours)?1:0;
    return {mask:mask,cols:cols,rows:rows,cellMm:cell,originX:originX,originY:originY,bounds:b};
  }
  function edt(mask, cols, rows) {
    var inf=1e20, tmp=new Float64Array(mask.length), out=new Float64Array(mask.length), max=Math.max(cols,rows), f=new Float64Array(max), d=new Float64Array(max), v=new Int32Array(max), z=new Float64Array(max+1);
    function pass(n) {var k=0;v[0]=0;z[0]=-inf;z[1]=inf;for(var q=1;q<n;q++){var s;do{s=((f[q]+q*q)-(f[v[k]]+v[k]*v[k]))/(2*q-2*v[k]);if(s<=z[k])k--;}while(s<=z[k]);k++;v[k]=q;z[k]=s;z[k+1]=inf;}k=0;for(q=0;q<n;q++){while(z[k+1]<q)k++;d[q]=(q-v[k])*(q-v[k])+f[v[k]];}}
    var x,y; for(y=0;y<rows;y++){for(x=0;x<cols;x++)f[x]=mask[y*cols+x]?inf:0;pass(cols);for(x=0;x<cols;x++)tmp[y*cols+x]=d[x];}
    for(x=0;x<cols;x++){for(y=0;y<rows;y++)f[y]=tmp[y*cols+x];pass(rows);for(y=0;y<rows;y++)out[y*cols+x]=Math.sqrt(d[y]);}
    return out;
  }
  function erode(mask, distance, clearanceCells) {var out=new Uint8Array(mask.length);for(var i=0;i<mask.length;i++)out[i]=mask[i]&&distance[i]+1e-9>=clearanceCells?1:0;return out;}
  function thin(input, cols, rows) {
    var a=new Uint8Array(input), remove=new Uint8Array(a.length), changed=true, guard=0;
    function phase(second){var any=false;remove.fill(0);for(var y=1;y<rows-1;y++)for(var x=1;x<cols-1;x++){var i=y*cols+x;if(!a[i])continue;var p2=a[i-cols],p3=a[i-cols+1],p4=a[i+1],p5=a[i+cols+1],p6=a[i+cols],p7=a[i+cols-1],p8=a[i-1],p9=a[i-cols-1],n=p2+p3+p4+p5+p6+p7+p8+p9;if(n<2||n>6)continue;var s=(!p2&&p3)+(!p3&&p4)+(!p4&&p5)+(!p5&&p6)+(!p6&&p7)+(!p7&&p8)+(!p8&&p9)+(!p9&&p2);if(s!==1)continue;if(second?(p2*p4*p8||p2*p6*p8):(p2*p4*p6||p4*p6*p8))continue;remove[i]=1;any=true;}for(var j=0;j<a.length;j++)if(remove[j])a[j]=0;return any;}
    while(changed&&guard++<Math.max(cols,rows)*2){changed=phase(false);if(phase(true))changed=true;}return a;
  }
  function neighbors(index, skeleton, cols, rows) {var y=Math.floor(index/cols),x=index-y*cols,out=[];for(var dy=-1;dy<=1;dy++)for(var dx=-1;dx<=1;dx++){if(!dx&&!dy)continue;var xx=x+dx,yy=y+dy;if(xx>=0&&yy>=0&&xx<cols&&yy<rows&&skeleton[yy*cols+xx])out.push(yy*cols+xx);}return out;}
  function trace(skeleton, grid, pruneMm) {
    var cols=grid.cols,rows=grid.rows, nodes=[], degree={}, visited={}, paths=[];
    for(var i=0;i<skeleton.length;i++)if(skeleton[i]){var n=neighbors(i,skeleton,cols,rows).length;degree[i]=n;if(n!==2)nodes.push(i);}
    function key(a,b){return a<b?a+':'+b:b+':'+a;} function pt(i){var y=Math.floor(i/cols),x=i-y*cols;return{x:grid.originX+(x+.5)*grid.cellMm,y:grid.originY+(y+.5)*grid.cellMm};}
    for(var ni=0;ni<nodes.length;ni++){var start=nodes[ni], ns=neighbors(start,skeleton,cols,rows);for(var q=0;q<ns.length;q++){if(visited[key(start,ns[q])])continue;var line=[pt(start)],prev=start,cur=ns[q];visited[key(prev,cur)]=1;while(true){line.push(pt(cur));if(degree[cur]!==2)break;var nn=neighbors(cur,skeleton,cols,rows),next=nn[0]===prev?nn[1]:nn[0];if(next===undefined||visited[key(cur,next)])break;prev=cur;cur=next;visited[key(prev,cur)]=1;}if(polylineLength(line)+1e-9>=pruneMm)paths.push(line);}}
    return paths;
  }
  function polylineLength(p){var l=0;for(var i=1;i<p.length;i++){var dx=p[i].x-p[i-1].x,dy=p[i].y-p[i-1].y;l+=Math.sqrt(dx*dx+dy*dy);}return l;}
  function sampleAt(p,d){for(var i=1;i<p.length;i++){var a=p[i-1],b=p[i],dx=b.x-a.x,dy=b.y-a.y,l=Math.sqrt(dx*dx+dy*dy);if(d<=l||i===p.length-1){var t=l?Math.min(1,d/l):0;return{x:a.x+dx*t,y:a.y+dy*t,angle:Math.atan2(dy,dx)};}d-=l;}return null;}
  function footprintFits(sample,w,h,contours,clearance){var ca=Math.cos(sample.angle),sa=Math.sin(sample.angle),hx=w/2+clearance,hy=h/2+clearance,pts=[],steps=8,i,t;for(i=0;i<=steps;i++){t=i/steps;pts.push([-hx+2*hx*t,-hy],[hx,-hy+2*hy*t],[hx-2*hx*t,hy],[-hx,hy-2*hy*t]);}pts.push([0,0]);for(i=0;i<pts.length;i++){var x=sample.x+pts[i][0]*ca-pts[i][1]*sa,y=sample.y+pts[i][0]*sa+pts[i][1]*ca;if(!pointInContours(x,y,contours))return false;}return true;}
  function place(path, contours, options){options=options||{};var L=polylineLength(path), inset=Math.max(0,Number(options.endpointInsetMm)||0), usable=L-2*inset,max=Math.max(.01,Number(options.maxSpacingMm)||50),out=[];if(usable<0)return out;var intervals=Math.max(1,Math.ceil(usable/max)),spacing=usable/intervals;for(var i=0;i<=intervals;i++){var s=sampleAt(path,inset+i*spacing);if(s&&footprintFits(s,Number(options.widthMm)||1,Number(options.heightMm)||1,contours,Number(options.clearanceMm)||0)){s.sequence=out.length+1;s.distanceMm=inset+i*spacing;out.push(s);}}out.actualSpacingMm=spacing;return out;}
  function generate(contours, options){options=options||{};var grid=rasterize(contours,options),dist=edt(grid.mask,grid.cols,grid.rows),clear=(Math.max(Number(options.widthMm)||0,Number(options.heightMm)||0)/2+(Number(options.clearanceMm)||0))/grid.cellMm,usable=erode(grid.mask,dist,clear),skel=thin(usable,grid.cols,grid.rows),guides=trace(skel,grid,Number(options.pruneMm)||grid.cellMm*2),placements=[];for(var i=0;i<guides.length;i++)placements.push(place(guides[i],contours,options));return{schemaVersion:1,grid:{cols:grid.cols,rows:grid.rows,cellMm:grid.cellMm,cells:grid.cols*grid.rows},guides:guides,placements:placements};}
  function splitSeries(modules,limit,wireReach){limit=Math.max(1,limit|0);var series=[],unreachable=[];for(var i=0;i<modules.length;i+=limit)series.push(modules.slice(i,i+limit));for(i=1;i<modules.length;i++){var dx=modules[i].x-modules[i-1].x,dy=modules[i].y-modules[i-1].y,gap=Math.sqrt(dx*dx+dy*dy);if(gap>wireReach)unreachable.push({after:i,gapMm:gap});}return{series:series,unreachable:unreachable};}
  return{pointInContours:pointInContours,rasterize:rasterize,distanceTransform:edt,thin:thin,trace:trace,polylineLength:polylineLength,placeModules:place,splitSeries:splitSeries,generate:generate};
}));
