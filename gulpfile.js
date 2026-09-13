const gulp = require('gulp');
const webpack = require('webpack');
const build = require('@microsoft/sp-build-web');
const getTasks = build.rig.getTasks;
const eslint = require('gulp-eslint-new');
const argv = require('yargs').argv;

build.tslintCmd.enabled = false;
build.rig.getTasks = function() {
  const result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

const eslintSubTask = build.subTask(
  'eslint',
  (gulp, buildOptions, done) => (
    gulp
      .src(['src/**/*.{ts,tsx}'])
      .pipe(eslint())
      .pipe(eslint.format())
      .pipe(eslint.failAfterError())
  )
);

build.configureWebpack.mergeConfig({
  additionalConfiguration(generatedConfig) {
    const config =  argv.config;
    const isProduction = (config !== undefined) && (config.toLowerCase() === 'prod');
    const isStaging = (config !== undefined) && (config.toLowerCase() === 'staging');
    const isDevelopment = (config === undefined) || (config.toLowerCase() === 'dev');
    const env = isProduction ? 'Prod' : isStaging ? 'Staging' : isDevelopment ? 'Dev' : '';
    let plugin, pluginDefine;
    for (let i = 0; i < generatedConfig.plugins.length; i++) {
      plugin = generatedConfig.plugins[i];
      if (plugin instanceof webpack.DefinePlugin) {
        pluginDefine = plugin;
      }
    }
    pluginDefine.definitions._ENVIRONMENT_ =  JSON.stringify(env);
    return generatedConfig;
  }
});



build.rig.addPreBuildTask(build.task('eslint-task', eslintSubTask));
/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(require('gulp'));
