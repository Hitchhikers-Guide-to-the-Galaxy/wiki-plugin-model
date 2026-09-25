import esbuild from 'esbuild'
import { writeFileSync } from 'node:fs'

const result = await esbuild.build({
  entryPoints: ['src/client/model.js'],
  bundle: true,
  format: 'iife',
  outfile: 'client/model.js',
  sourcemap: true,
  minify: true,
  metafile: true,
})
writeFileSync('meta-client.json', JSON.stringify(result.metafile))
console.log('built client/model.js')
