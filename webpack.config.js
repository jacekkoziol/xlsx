
const path = require('path');
const HtmlWebpackPlugin = require("html-webpack-plugin");
const { CleanWebpackPlugin } = require('clean-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: './src/index.ts',
    devtool: 'inline-source-map',
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: 'ts-loader',
          exclude: /node_modules/
        },
      ]
    },
    resolve: {
      extensions: ['.ts', '.js', '.tsx']
    },
    output: {
      filename: 'bundle.js',
      path: path.resolve(__dirname, 'dist')
    },
    plugins: [
      new CleanWebpackPlugin({ cleanStaleWebpackAssets: false }),
      new HtmlWebpackPlugin({
        title: "Webpack Output",
        template: './src/index.html'
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: './src/xlsx-templates', to: 'xlsx-templates' },
          { from: './src/styles' },
        ]
      })
    ],
    devServer: {
      contentBase: './dist',
    },
};