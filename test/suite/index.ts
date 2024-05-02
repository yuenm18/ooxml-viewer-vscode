import { glob } from 'glob';
import Mocha from 'mocha';
import path from 'path';

export async function run(): Promise<void> {
  // Create the mocha test
  const mocha = new Mocha({
    ui: 'tdd',
  });
  // mocha.useColors(true);

  const testsRoot = path.resolve(__dirname, '..');

  const files = await glob.glob('**/**.test.js', { cwd: testsRoot });

  // Add files to the test suite
  files.forEach(f => mocha.addFile(path.resolve(testsRoot, f)));

  // Run the mocha test

  return new Promise((c, e) => {
    mocha.run(failures => {
      if (failures > 0) {
        e(new Error(`${failures} tests failed.`));
      } else {
        c();
      }
    });
  });
}
