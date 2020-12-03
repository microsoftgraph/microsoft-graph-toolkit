const gulp = require('gulp');
const sass = require('gulp-sass');
const gap = require('gulp-append-prepend');
const cleanCSS = require('gulp-clean-css');
const rename = require('gulp-rename');
var license = require('gulp-header-license');

scssFileHeader = `
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD
// MODIFY THE .SCSS FILE INSTEAD

import { css } from 'lit-element';
/**
 * exports lit-element css
 * @export styles
 */
export const styles = [
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

function runSass() {
  return gulp
    .src('src/**/!(shared)*.scss')
    .pipe(sass())
    .pipe(cleanCSS())
    .pipe(gap.prependText(scssFileHeader))
    .pipe(gap.appendText(scssFileFooter))
    .pipe(rename({ extname: '-css.ts' }))
    .pipe(gulp.dest('src/'));
}

function setLicense() {
  return gulp
    .src(['src/**/*.{ts,js,scss}', 'tests/**/*.{ts,js,scss}'])
    .pipe(license(licenseStr))
    .pipe(gulp.dest('src/'));
}

function setVersion() {
  var pkg = require(process.cwd() + '/package.json');
  var fs = require('fs');
  fs.writeFileSync('./src/utils/version.ts', versionFile.replace('[VERSION]', pkg.version));
}

gulp.task('sass', runSass);
gulp.task('setLicense', setLicense);
gulp.task('setVersion', async () => setVersion());

gulp.task('watchSass', () => {
  runSass();
  return gulp.watch('src/**/*.scss', gulp.series('sass'));
});
