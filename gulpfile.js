const gulp = require("gulp");
const sass = require("gulp-sass");
const gap = require("gulp-append-prepend");
const rename = require("gulp-rename");

const header = "export default `";
const footer = "`;";

function processSass() {
  return gulp
    .src("src/**/*.scss")
    .pipe(sass())
    .pipe(gap.prependText(header))
    .pipe(gap.appendText(footer))
    .pipe(rename({ extname: ".ts" }));
}

gulp.task("sass", processSass);
gulp.task("sass:watch", () => {
  processSass();
  return gulp.watch("src/**/*.scss", gulp.series("sass"));
});
