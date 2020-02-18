const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require('webpack');

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      vendor: [
        'react',
        'react-dom',
        'core-js/stable',
        'regenerator-runtime/runtime',
        'office-ui-fabric-react'
      ],
      taskpane: [
        'react-hot-loader/patch',
        './src/taskpane/index.tsx',
      ],
      //commands: './src/commands/commands.ts',
      login: './src/login/login.tsx',
      logout: './src/logout/logout.tsx'
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: [
            'react-hot-loader/webpack',
            'ts-loader'
          ],
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          use: {
            loader: 'file-loader',
            query: {
              name: 'assets/[name].[ext]'
            }
          }
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new CopyWebpackPlugin([
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        }
      ]),
      new MiniCssExtractPlugin('[name].[hash].css'),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane', 'vendor', 'polyfills']
      }),
      /*new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"]
      }),*/
      new HtmlWebpackPlugin({
        filename: "login.html",
        template: "./src/login/login.html",
        chunks: ["login"]
      }),
      new HtmlWebpackPlugin({
        filename: "logout.html",
        template: "./src/logout/logout.html",
        chunks: ["logout"]
      }),
      new CopyWebpackPlugin([
        {
          from: './assets',
          ignore: ['*.scss'],
          to: 'assets',
        }
      ]),
      new CopyWebpackPlugin([
        {
          from: './src/auth.html',
          to: 'auth.html',
        }
      ]),
      new CopyWebpackPlugin([
        {
          from: './src/index.html',
          to: 'index.html',
        }
      ]),
      new CopyWebpackPlugin([
        {
          from: './manifest.xml',
          to: 'manifest.xml',
        }
      ]),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};
