'use strict'

series = require("gulp").series
parallel = require("gulp").parallel
watch = require("gulp").watch
task = require("gulp").task


mainTask = require("./main.coffee")




task "default", mainTask


task "bot", (cb)->

  watch mainTask.watch, mainTask

