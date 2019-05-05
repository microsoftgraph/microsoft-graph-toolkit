const gulp = require('gulp');
const sass = require('gulp-sass');
const gap = require('gulp-append-prepend');
const rename = require('gulp-rename');

scssFileHeader = `
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD
// MODIFY THE .SCSS FILE INSTEAD

import { css } from 'lit-element';
export const styles = [
  css\``;

scssFileFooter = '`];';

versionFile = `/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD

export const PACKAGE_VERSION = '[VERSION]';
`;

function runSass() {
  return gulp
    .src('src/**/!(shared-styles).scss')
    .pipe(sass())
    .pipe(gap.prependText(scssFileHeader))
    .pipe(gap.appendText(scssFileFooter))
    .pipe(rename({ extname: '-css.ts' }))
    .pipe(gulp.dest('src/'));
}

function setVersion() {
  var pkg = require('./package.json');
  var fs = require('fs');
  fs.writeFileSync('src/utils/version.ts', versionFile.replace('[VERSION]', pkg.version));
}

gulp.task('sass', runSass);
gulp.task('setVersion', async () => setVersion());

gulp.task('watchSass', () => {
  runSass();
  return gulp.watch('src/**/*.scss', gulp.series('sass'));
});
