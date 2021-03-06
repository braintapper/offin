gulp   = require('gulp')
coffee = require("gulp-coffee")

# set the compilation order of the client
sourcePaths = ["src/**/*.coffee"]

module.exports = (cb)->
  destinationPath = "./"



  stream = gulp.src(sourcePaths)
  .pipe(coffee({bare:true}))
  .pipe(gulp.dest(destinationPath)) #.pipe(jsmin()).pipe(gulp.dest(destinationPath))
  stream.on 'end', ()->
    cb()

module.exports.watch = sourcePaths
module.exports.displayName = "main"
