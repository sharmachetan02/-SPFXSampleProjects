/*
* User webpack settings file. You can add your own settings here.
* Changes from this file will be merged into the base webpack configuration file.
* This file will not be overwritten by the subsequent spfx-fast-serve calls.
*/
const webpack = require("webpack");
const argv = require("yargs").argv;
/**
 * you can add your project related webpack configuration here, it will be merged using webpack-merge module
 * i.e. plugins: [new webpack.Plugin()]
 */
const webpackConfig = {

}

/**
 * For even more fine-grained control, you can apply custom webpack settings using below function
 * @param {object} initialWebpackConfig - initial webpack config object
 * @param {object} webpack - webpack object, used by SPFx pipeline
 * @returns webpack config object
 */
const transformConfig = function (initialWebpackConfig, webpack) {
  // transform the initial webpack config here, i.e.
  // initialWebpackConfig.plugins.push(new webpack.Plugin()); etc.
  const configArg = argv.config.config;
  const isProduction = configArg && configArg.toLowerCase() === 'prod';
  const isStaging = configArg && configArg.toLowerCase() === 'staging';
  const isDevelopment = !configArg || configArg.toLowerCase() === 'dev';

  const env = isProduction ? 'Prod' : isStaging ? 'Staging' : isDevelopment ? 'Dev' : '';

  // 🔄 Remove existing DefinePlugin (if any)
  initialWebpackConfig.plugins = initialWebpackConfig.plugins.filter(
    (plugin) => !(plugin instanceof webpack.DefinePlugin)
  );

  // ✅ Add DefinePlugin manually
  initialWebpackConfig.plugins.push(
    new webpack.DefinePlugin({
      _ENVIRONMENT_: JSON.stringify(env)
    })
  );
  return initialWebpackConfig;
}

module.exports = {
  webpackConfig,
  transformConfig
}
