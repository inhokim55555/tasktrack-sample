// Refuse ambient configuration files; samples never read production credentials.
import {existsSync} from 'node:fs';import {fileURLToPath} from 'node:url';
for(const f of ['.env','.env.local','.env.production','.env.production.local','.env.development','.env.development.local'])if(existsSync(fileURLToPath(new URL('../'+f,import.meta.url))))throw Error('Remove '+f+' from this sample folder. No credentials are needed.');
