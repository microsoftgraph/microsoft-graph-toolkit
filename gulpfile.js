const gulp = require('gulp');
const sass = require('gulp-sass');
const gap = require('gulp-append-prepend');
const rename = require('gulp-rename');


scssFileHeader =`
// THIS FILE IS AUTO GENERATED
// ANY CHANGES WILL BE LOST DURING BUILD
// MODIFY THE .SCSS FILE INSTEAD

import { css } from 'lit-element';
import { sharedStyles } from '../../styles/shared-styles';
export const styles = [
  sharedStyles,
  css\``

scssFileFooter = '`];';

function runSass() {
  return gulp
    .src('src/**/*.scss')
    .pipe(sass())
    .pipe(gap.prependText(scssFileHeader))
    .pipe(gap.appendText(scssFileFooter))
    .pipe(rename({ extname: '.scss.ts' }))
    .pipe(gulp.dest('src/'));
}

gulp.task('sass', runSass);

gulp.task('watchSass', () => {
  return gulp.watch('src/**/*.scss', gulp.series('sass'));
});
