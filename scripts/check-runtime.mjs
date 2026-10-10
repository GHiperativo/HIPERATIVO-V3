const requiredMajor = 24;
const nodeVersion = process.versions.node;
const major = Number(nodeVersion.split('.')[0]);

console.log(`Node.js: v${nodeVersion}`);
console.log(`Platform: ${process.platform}/${process.arch}`);
console.log(`npm: ${process.env.npm_config_user_agent ?? 'não detectado fora do npm'}`);

if (major !== requiredMajor) {
  console.error(`Runtime inválido: esperado Node.js ${requiredMajor}.x, recebido ${nodeVersion}.`);
  process.exit(1);
}

console.log('Runtime Node.js validado com sucesso.');
