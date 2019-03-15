const gulp = require('gulp');
const sass = require('gulp-sass');
const gap = require('gulp-append-prepend');
const rename = require('gulp-rename');

function runSass() {
  return gulp
    .src('src/**/*.scss')
    .pipe(sass())
    .pipe(
      gap.prependText(
        `
        import { css } from 'lit-element';
        import { sharedStyles } from '../../styles/shared-styles';
        export const styles = [
          sharedStyles,
          css\`
            @import 'styles/fabric-styles.css';
        `
      )
    )
    .pipe(gap.appendText('`];'))
    .pipe(rename({ extname: '-styles.ts' }))
    .pipe(gulp.dest('src/'));
}

gulp.task('sass', runSass);

gulp.task('watchSass', () => {
  return gulp.watch('src/**/*.scss', gulp.series('sass'));
});
