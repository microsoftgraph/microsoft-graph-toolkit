const gulp = require('gulp');
const license = require('gulp-header-license');

scssFileHeader = `
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD
// MODIFY THE .SCSS FILE INSTEAD

import { css, CSSResult } from 'lit';
/**
 * exports lit-element css
 * @export styles
 */
export const styles: CSSResult[] = [
  css\``;

scssFileFooter = '`];';

licenseStr = `/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */
`;

versionFile = `${licenseStr}
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD

export const PACKAGE_VERSION = '[VERSION]';
`;

function setLicense() {
  return gulp
    .src(['packages/**/src/**/*.{ts,js,scss}', '!packages/**/generated/**/*'], { base: './' })
    .pipe(license(licenseStr))
    .pipe(gulp.dest('./'));
}

function setVersion() {
  var pkg = require('./package.json');
  var fs = require('fs');
  fs.writeFileSync('./src/utils/version.ts', versionFile.replace('[VERSION]', pkg.version));
}

gulp.task('setLicense', setLicense);
gulp.task('setVersion', async () => setVersion());
