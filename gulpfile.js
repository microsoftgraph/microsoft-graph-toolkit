const gulp = require('gulp');
const sass = require('gulp-sass')(require('sass'));
const gap = require('gulp-append-prepend');
const cleanCSS = require('gulp-clean-css');
const rename = require('gulp-rename');
const license = require('gulp-header-license');
const replace = require('gulp-replace-task');

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

function runSass() {
  return (
    gulp
      .src('src/**/!(shared)*.scss')
      .pipe(sass())
      .pipe(cleanCSS())
      // replacement to make office-ui-fabric-core icons work with lit-element
      .pipe(
        replace({
          patterns: [
            {
              match: /:"\\([0-9a-f])/g,
              replacement: ':"\\u$1'
            }
          ]
        })
      )
      .pipe(gap.prependText(scssFileHeader))
      .pipe(gap.appendText(scssFileFooter))
      .pipe(rename({ extname: '-css.ts' }))
      .pipe(gulp.dest('src/'))
  );
}

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

gulp.task('sass', runSass);
gulp.task('setLicense', setLicense);
gulp.task('setVersion', async () => setVersion());

gulp.task('watchSass', () => {
  runSass();
  return gulp.watch('src/**/*.scss', gulp.series('sass'));
});
