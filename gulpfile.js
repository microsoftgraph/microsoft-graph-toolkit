const gulp = require('gulp');
const sass = require('gulp-sass');
const gap = require('gulp-append-prepend');
const rename = require('gulp-rename');

function runSass() {
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
}

gulp.task('sass', runSass);

gulp.task('watchSass', () => {
    return gulp.watch('src/**/*.scss', gulp.series('sass'));
})