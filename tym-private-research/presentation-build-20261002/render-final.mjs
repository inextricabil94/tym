import fs from 'node:fs/promises';
import path from 'node:path';
import { FileBlob, PresentationFile } from '@oai/artifact-tool';

const finalPath=process.argv[2]??'C:/Users/serba/Documents/ChatGPT/PhD/output/TYM_MLNET_Updated_Results_20261002.pptx';
const prefix=process.argv[3]??'final';
const privateDir='C:/Users/serba/Documents/ChatGPT/PhD/tym-private-research/presentation-build-20261002';
const p=await PresentationFile.importPptx(await FileBlob.load(finalPath));
for(let i=0;i<p.slides.items.length;i++) {
 const png=await p.export({slide:p.slides.items[i],format:'png',scale:1});
 await fs.writeFile(path.join(privateDir,`${prefix}-slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await png.arrayBuffer()));
}
const montage=await p.export({format:'webp',montage:true,scale:0.4});
await fs.writeFile(path.join(privateDir,`${prefix}-contact-sheet.webp`),new Uint8Array(await montage.arrayBuffer()));
const snap=await p.inspect({kind:'slide,table,chart,notes',maxChars:80000});
await fs.writeFile(path.join(privateDir,`${prefix}-inspection.ndjson`),snap.ndjson);
console.log(JSON.stringify({slides:p.slides.items.length,finalRendered:true,privateDir}));
