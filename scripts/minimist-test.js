var argv = require('minimist')(process.argv.slice(2));
console.dir(argv.a);
// command: node minimist-test -a beep -b boop
// outputs: { _: [], a: 'beep', b: 'boop' }
