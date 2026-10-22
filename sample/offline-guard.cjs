// Local samples never contact external services, even when ambient credentials exist.
const net=require('node:net');const {syncBuiltinESMExports}=require('node:module');
const allowed=h=>['localhost','127.0.0.1','::1','[::1]'].includes(String(h||'localhost').toLowerCase());
const connect=net.Socket.prototype.connect;
net.Socket.prototype.connect=function(...args){let o=args[0];if(Array.isArray(o))o=o[0];if(o&&typeof o==='object'){if(o.path)throw Error('Socket paths disabled in sample');if(!allowed(o.host))throw Error('External connections disabled in sample');}else if(typeof args[1]==='string'&&!allowed(args[1]))throw Error('External connections disabled in sample');return connect.apply(this,args);};
const nativeFetch=globalThis.fetch;globalThis.fetch=async(input,...rest)=>{const u=new URL(typeof input==='string'||input instanceof URL?input:input.url);if(!allowed(u.hostname))throw Error('External requests disabled in sample');return nativeFetch(input,...rest);};syncBuiltinESMExports();
if(process.send)process.on('disconnect',()=>process.exit(0));if(process.channel)process.channel.unref();
