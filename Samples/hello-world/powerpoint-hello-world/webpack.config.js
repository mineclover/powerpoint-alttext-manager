const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
  const isProduction = argv.mode === 'production';

  return {
    entry: {
      taskpane: './src/taskpane/taskpane.ts'
    },
    output: {
      filename: '[name].js',
      path: path.resolve(__dirname, 'dist'),
      clean: true
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.js', '.css']
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: 'ts-loader',
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane']
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets',
            to: 'assets',
            noErrorOnMissing: true
          },
          {
            from: 'manifest*.xml',
            to: '[name][ext]'
          }
        ]
      })
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, 'dist')
      },
      server: {
        type: 'https',
        options: {
          key: path.resolve(__dirname, 'localhost.key'),
          cert: path.resolve(__dirname, 'localhost.crt')
        }
      },
      port: 3000,
      hot: true,
      open: false,
      headers: {
        'Access-Control-Allow-Origin': '*'
      }
    },
    devtool: isProduction ? 'source-map' : 'eval-source-map',
    mode: argv.mode || 'development'
  };
};
