// fixes.js
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const zipFileName = process.argv[2];

if (!zipFileName) {
  console.error('Usage: node fixes.js <zip-file-name>');
  console.error('Example: node fixes.js scoped-edit-assistant-flat.zip');
  process.exit(1);
}

const DOWNLOADS_DIR = 'C:\\Users\\lngo\\Downloads';
const ZIP_PATH = path.join(DOWNLOADS_DIR, zipFileName);
const PROJECT_ROOT = __dirname;
const TEMP_EXTRACT_DIR = path.join(PROJECT_ROOT, '.temp_extracted_fixes');
const TIMESTAMP = new Date().toISOString().replace(/[:.]/g, '-');
const BACKUP_ROOT = path.join(PROJECT_ROOT, '.backup', TIMESTAMP);
const SUBDIRS_TO_SYNC = [''];
const KNOWN_PROJECT_FILES = [
  'taskpane.html',
  'taskpane.js',
  'styles.css',
  'commands.html',
  'commands.js',
  'README.txt',
  'manifest.xml'
];

function exists(p) {
  return fs.existsSync(p);
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function psEscape(value) {
  return value.replace(/'/g, "''");
}

function expandZip(zipPath, destinationPath) {
  ensureDir(destinationPath);
  execFileSync(
    'powershell.exe',
    [
      '-NoProfile',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      `Expand-Archive -LiteralPath '${psEscape(zipPath)}' -DestinationPath '${psEscape(destinationPath)}' -Force`
    ],
    { stdio: 'inherit' }
  );
}

function walkDirs(startDir, out = []) {
  const entries = fs.readdirSync(startDir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(startDir, entry.name);
    if (entry.isDirectory()) {
      out.push(fullPath);
      walkDirs(fullPath, out);
    }
  }
  return out;
}

function containsKnownProjectFiles(dir) {
  if (!exists(dir) || !fs.statSync(dir).isDirectory()) return false;
  const entries = new Set(fs.readdirSync(dir));
  return KNOWN_PROJECT_FILES.some(name => entries.has(name));
}

function findExtractedProjectRoot(startDir) {
  if (containsKnownProjectFiles(startDir)) {
    return startDir;
  }

  const dirs = walkDirs(startDir);
  for (const dir of dirs) {
    if (containsKnownProjectFiles(dir)) {
      return dir;
    }
  }

  const topLevelDirs = fs
    .readdirSync(startDir, { withFileTypes: true })
    .filter(entry => entry.isDirectory())
    .map(entry => path.join(startDir, entry.name));

  if (topLevelDirs.length === 1) {
    return topLevelDirs[0];
  }

  throw new Error(
    `Could not find extracted project root containing known project files under: ${startDir}`
  );
}

function collectFiles(dir, baseDir = dir, out = []) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      collectFiles(fullPath, baseDir, out);
    } else if (entry.isFile()) {
      out.push({
        fullPath,
        relativePath: path.relative(baseDir, fullPath)
      });
    }
  }
  return out;
}

function filesAreEqual(fileA, fileB) {
  if (!exists(fileA) || !exists(fileB)) return false;
  const statA = fs.statSync(fileA);
  const statB = fs.statSync(fileB);
  if (statA.size !== statB.size) return false;
  const bufA = fs.readFileSync(fileA);
  const bufB = fs.readFileSync(fileB);
  return bufA.equals(bufB);
}

function copyFileWithBackup(sourceFile, destFile) {
  const result = {
    action: 'skipped',
    sourceFile,
    destFile,
    backupFile: null,
    error: null
  };

  if (!exists(sourceFile)) {
    result.action = 'error';
    result.error = `Source file missing: ${sourceFile}`;
    return result;
  }

  if (!exists(destFile)) {
    result.action = 'error';
    result.error = `Destination file does not already exist in project: ${destFile}`;
    return result;
  }

  if (filesAreEqual(sourceFile, destFile)) {
    return result;
  }

  const relativeToProject = path.relative(PROJECT_ROOT, destFile);
  const backupFile = path.join(BACKUP_ROOT, relativeToProject);
  ensureDir(path.dirname(backupFile));
  fs.copyFileSync(destFile, backupFile);
  result.backupFile = backupFile;

  ensureDir(path.dirname(destFile));
  fs.copyFileSync(sourceFile, destFile);
  result.action = 'updated';
  return result;
}

function cleanupTempDir() {
  if (exists(TEMP_EXTRACT_DIR)) {
    fs.rmSync(TEMP_EXTRACT_DIR, { recursive: true, force: true });
  }
}

function main() {
  if (!exists(ZIP_PATH)) {
    throw new Error(`ZIP file not found: ${ZIP_PATH}`);
  }

  if (!exists(PROJECT_ROOT)) {
    throw new Error(`Project root not found: ${PROJECT_ROOT}`);
  }

  cleanupTempDir();

  console.log(`\nExtracting ZIP: ${ZIP_PATH}`);
  expandZip(ZIP_PATH, TEMP_EXTRACT_DIR);

  const extractedProjectRoot = findExtractedProjectRoot(TEMP_EXTRACT_DIR);
  console.log(`Found extracted project root: ${extractedProjectRoot}`);

  const results = [];

  for (const subdir of SUBDIRS_TO_SYNC) {
    const sourceSubdir = path.join(extractedProjectRoot, subdir);
    const destSubdir = path.join(PROJECT_ROOT, subdir);

    if (!exists(sourceSubdir)) {
      results.push({ action: 'error', error: `Missing extracted folder: ${sourceSubdir}` });
      continue;
    }

    if (!exists(destSubdir)) {
      results.push({ action: 'error', error: `Missing destination folder in project: ${destSubdir}` });
      continue;
    }

    const sourceFiles = collectFiles(sourceSubdir);

    for (const file of sourceFiles) {
      const sourceFile = file.fullPath;
      const destFile = path.join(destSubdir, file.relativePath);
      const result = copyFileWithBackup(sourceFile, destFile);
      results.push(result);

      const relDest = path.relative(PROJECT_ROOT, destFile);
      if (result.action === 'updated') {
        const relBackup = path.relative(PROJECT_ROOT, result.backupFile);
        console.log(`UPDATED  ${relDest}`);
        console.log(`BACKUP   ${relBackup}`);
      } else if (result.action === 'skipped') {
        console.log(`SKIPPED  ${relDest} (no changes)`);
      } else if (result.action === 'error') {
        console.error(`ERROR    ${relDest}`);
        console.error(`         ${result.error}`);
      }
    }
  }

  const updatedCount = results.filter(r => r.action === 'updated').length;
  const skippedCount = results.filter(r => r.action === 'skipped').length;
  const errorResults = results.filter(r => r.action === 'error');

  console.log('\nDone.');
  console.log(`Updated: ${updatedCount}`);
  console.log(`Skipped: ${skippedCount}`);
  console.log(`Errors: ${errorResults.length}`);

  if (updatedCount > 0) {
    console.log(`Backups saved to: ${BACKUP_ROOT}`);
  } else {
    console.log('No backups were needed.');
  }

  if (errorResults.length > 0) {
    console.log('\nError summary:');
    for (const err of errorResults) {
      console.log(`- ${err.error}`);
    }
    process.exitCode = 1;
  }
}

try {
  main();
} catch (error) {
  console.error('\nERROR:', error.message);
  process.exitCode = 1;
} finally {
  cleanupTempDir();
}