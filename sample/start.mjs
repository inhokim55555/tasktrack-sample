import './configure.mjs';
import http from 'node:http';import {fork} from 'node:child_process';import {readFile,realpath,stat,mkdir,copyFile} from 'node:fs/promises';import path from 'node:path';import {fileURLToPath} from 'node:url';
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..'),state=path.join(root,'sample-state');
const PORT=4296,BACKEND=0,DBPORT=0,id="tasktrack";const children=[];
await mkdir(state,{recursive:true});
const env={};for(const k of ['SystemRoot','WINDIR','TEMP','TMP','PATH','COMSPEC','PATHEXT','LOCALAPPDATA'])if(process.env[k])env[k]=process.env[k];
Object.assign(env,{NODE_ENV:id==='trackify'?'production':'development',PORT:String(BACKEND),HOSTNAME:'127.0.0.1',HOST:'127.0.0.1',PORTFOLIO_SAMPLE:'1',NEXT_PUBLIC_PORTFOLIO_SAMPLE:'1',NEXT_TELEMETRY_DISABLED:'1',DATA_SOURCE:'mock',DATABASE_URL:'postgresql://postgres:postgres@127.0.0.1:'+DBPORT+'/postgres?connection_limit=1&pool_timeout=30&pgbouncer=true&statement_cache_size=0',BYOU_PNL_DATABASE_URL:'postgresql://postgres:postgres@127.0.0.1:'+DBPORT+'/postgres',BYOU_PNL_DB_POOL_SIZE:'1',BYOU_DATA_DIR:path.join(state,'intel'),BYOU_ALLOW_AUTH_DISABLED:'true',BYOU_SESSION_SECRET:'public-fictional-local-sample',BYOU_ADMIN_TOKEN:'sample-only',BYOU_INVENTORY_APPROVAL_PASSWORD:'sample'});
function child(file,args=[]){const p=fork(path.join(root,file),args,{cwd:root,env,execArgv:['--require',path.join(root,'sample/offline-guard.cjs')],stdio:['ignore','inherit','inherit','ipc'],windowsHide:true});children.push(p);return p;}
const delay=ms=>new Promise(r=>setTimeout(r,ms));
async function run(file){const p=child(file);await new Promise((r,j)=>{p.once('error',j);p.once('exit',n=>n===0?r():j(Error(file+' failed')));});}
async function stop(){for(const p of children)p.kill();server?.close();setTimeout(()=>process.exit(0),300).unref();}let server;
process.on('SIGINT',stop);process.on('SIGTERM',stop);
if(id!=='tasktrack'){
 const db=child('sample/db-host.mjs',[id==='trackify'?'trackify':'intel',path.join(state,'database'),String(DBPORT)]);
 await new Promise((r,j)=>{db.once('message',r);db.once('error',j);db.once('exit',()=>j(Error('Database failed to start')));});
 if(id==='trackify'){await run('scripts/portfolio-seed.mjs');child('node_modules/next/dist/bin/next',['start','--hostname','127.0.0.1','--port',String(BACKEND)]);}
 else {await mkdir(env.BYOU_DATA_DIR,{recursive:true});for(const f of ['latest-dataset.json','latest-startup.json']){try{await stat(path.join(env.BYOU_DATA_DIR,f));}catch{await copyFile(path.join(root,'sample/seed',f),path.join(env.BYOU_DATA_DIR,f));}}await run('sample/seed-pnl.mjs');child('server.mjs');}
 let ready=false;for(let i=0;i<100;i++){try{const r=await fetch('http://127.0.0.1:'+BACKEND,{signal:AbortSignal.timeout(1000)});if(r.status<500){ready=true;break;}}catch{}await delay(150);}if(!ready){await stop();throw Error('Sample backend failed to start');}
}
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript','.css':'text/css','.json':'application/json','.svg':'image/svg+xml','.png':'image/png','.woff2':'font/woff2','.ttf':'font/ttf','.ico':'image/x-icon'};
server=http.createServer(async(req,res)=>{
 const fail=(n,s)=>{res.writeHead(n);res.end(s);};
 res.setHeader('Content-Security-Policy',"default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'; frame-src 'none'; object-src 'none'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'");res.setHeader('X-Content-Type-Options','nosniff');res.setHeader('Referrer-Policy','no-referrer');res.setHeader('Cache-Control','no-store');
 if(req.headers.host!=='127.0.0.1:'+PORT&&req.headers.host!=='localhost:'+PORT)return fail(403,'Local host required');
 if(!['GET','HEAD'].includes(req.method)&&req.headers.origin!=='http://127.0.0.1:'+PORT&&req.headers.origin!=='http://localhost:'+PORT)return fail(403,'Local origin required');
 try{const u=new URL(req.url,'http://127.0.0.1');const pathname=decodeURIComponent(u.pathname);if(pathname.includes('\\')||pathname.split('/').some(p=>p.startsWith('.')))return fail(404,'Not found');
 if(pathname==='/sample-guard.js'){res.setHeader('Content-Type','text/javascript');return res.end(await readFile(path.join(root,'public/sample-guard.js')));}
 if(BACKEND){const headers={...req.headers,host:'127.0.0.1:'+BACKEND};delete headers['x-forwarded-host'];delete headers['x-forwarded-proto'];headers['accept-encoding']='identity';if(headers.origin)headers.origin='http://127.0.0.1:'+BACKEND;
 const p=http.request({hostname:'127.0.0.1',port:BACKEND,path:req.url,method:req.method,headers},up=>{res.statusCode=up.statusCode;for(const [k,v]of Object.entries(up.headers))if(v!==undefined&&!['content-security-policy','x-frame-options','content-length','content-encoding'].includes(k))res.setHeader(k,v);if(String(up.headers['content-type']).includes('text/html')){const chunks=[];up.on('data',c=>chunks.push(c));up.on('end',()=>res.end(Buffer.concat(chunks).toString().replace('</head>','<script src="/sample-guard.js"></script></head>')));}else up.pipe(res);});p.on('error',()=>{if(!res.headersSent)fail(502,'Sample unavailable');else res.end();});req.pipe(p);return;}
 if(!['GET','HEAD'].includes(req.method))return fail(405,'Read only');const folder=await realpath(path.join(root,'dist'));const candidate=await realpath(path.resolve(folder,'.'+(pathname==='/'?'/index.html':pathname)));if(!candidate.startsWith(folder+path.sep)||!(await stat(candidate)).isFile())return fail(404,'Not found');const type=types[path.extname(candidate)];if(!type)return fail(404,'Not found');res.setHeader('Content-Type',type);res.end(req.method==='HEAD'?undefined:await readFile(candidate));
 }catch{if(!res.headersSent)fail(404,'Not found');else res.end();}
});server.listen(PORT,'127.0.0.1',()=>console.log('Sample ready: http://127.0.0.1:'+PORT));
