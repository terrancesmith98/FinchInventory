/// <binding BeforeBuild='compileSass' ProjectOpened='watch' />
var gulp = require('gulp'),
    sass = require('gulp-sass'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename'),
    maps = require('gulp-sourcemaps');

gulp.task('compileSass', function () {
    return gulp.src('./src/scss/**/*.scss')
        .pipe(maps.init())
        .pipe(sass().on('error', sass.logError))
        .pipe(maps.write())
        .pipe(gulp.dest('./Content'));
});

//gulp.task('concatScripts', function () {
//    return gulp.src([
//        'src/js/main.js'
//    ])
//        .pipe(concat('app.js'))
//        .pipe(maps.init())
//        .pipe(maps.write())
//        .pipe(gulp.dest('Scripts'));
//});

//gulp.task('minifyScripts', ['concatScripts'], function () {
//    return gulp.src('Scripts/app.js')
//        .pipe(uglify({ mangle: false }))
//        .pipe(rename('app.min.js'))
//        .pipe(gulp.dest('Scripts'));
//});


gulp.task('watch', function () {
    gulp.watch('./src/scss/**/*.scss', gulp.series('compileSass'));
    //gulp.watch('./src/js/**/*.js', ['minifyScripts']);
});

gulp.task('default', function () {
    //gulp.task('minifyScripts', ['compileSass']);
});