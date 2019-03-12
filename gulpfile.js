const gulp = require('gulp');
const sass = require('gulp-sass');
const gap = require('gulp-append-prepend');
const rename = require('gulp-rename');

gulp.task('sass', () => {
    return gulp.src('src/**/*.scss')
        .pipe(sass())
        .pipe(gap.prependText(`import { css } from 'lit-element';
import { globalStyle } from '../../global/variables.style';
export const style = [
    globalStyle,
    css\``))
        .pipe(gap.appendText('`];'))
        .pipe(rename({extname: ".style.ts"}))
        .pipe(gulp.dest('src/'));
});