import test from 'node:test';
import assert from 'node:assert/strict';
import {createRequire} from 'node:module';
import net from 'node:net';
createRequire(import.meta.url)('./offline-guard.cjs');
test('external fetch is rejected before any network connection',async()=>{await assert.rejects(fetch('https://example.invalid/private'),/External requests disabled/);});
test('external sockets and local filesystem sockets are rejected',()=>{assert.throws(()=>new net.Socket().connect({host:'example.invalid',port:443}),/External connections disabled/);assert.throws(()=>new net.Socket().connect({path:'private.sock'}),/Socket paths disabled/);});
