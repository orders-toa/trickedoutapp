const { spawn } = require('child_process');
const readline = require('readline');

const VERSION_REGEX =
  /^\d+\.\d+\.\d+(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$/;

function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      stdio: 'inherit',
      shell: false,
      windowsHide: true,
    });

    child.on('exit', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${command} ${args.join(' ')} exited with code ${code}`));
    });
    child.on('error', reject);
  });
}

function ask(question) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve) => {
    rl.question(question, (answer) => {
      rl.close();
      resolve(answer.trim());
    });
  });
}

async function main() {
  if (!process.env.GH_TOKEN) {
    console.error('GH_TOKEN is required to publish releases to GitHub.');
    console.error('PowerShell: $env:GH_TOKEN="your_token"');
    console.error('CMD: set GH_TOKEN=your_token');
    process.exit(1);
  }

  const version = await ask('Enter release version (example: 1.0.1): ');
  if (!VERSION_REGEX.test(version)) {
    console.error('Invalid version format. Use semver, for example 1.0.1 or 1.1.0-beta.1');
    process.exit(1);
  }

  const npmCmd = process.platform === 'win32' ? 'npm.cmd' : 'npm';
  const npxCmd = process.platform === 'win32' ? 'npx.cmd' : 'npx';

  try {
    console.log(`\nUpdating package version to ${version}...`);
    await run(npmCmd, ['version', version, '--no-git-tag-version']);

    console.log('\nPublishing release to GitHub...');
    await run(npxCmd, ['electron-builder', '--win', '--x64', '--publish', 'always']);

    console.log('\nRelease complete.');
  } catch (err) {
    console.error(`\nRelease failed: ${err.message}`);
    process.exit(1);
  }
}

main();
