/* eslint-disable no-undef */
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const path = require("path");

// Cambia esto si vas a publicar en producción
const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/";

/* global require, module, process, __dirname */

module.exports = async (_env, options) => {
  const dev = options.mode === "development";

  return {
    devtool: "source-map",

    // 👇 Solo las entradas que realmente usamos (no hay 'functions' aquí)
    entry: {
      polyfill: "core-js/stable",
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
    },

    output: {
      clean: true,
    },

    resolve: {
      extensions: [".html", ".js"],
    },

    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: { loader: "babel-loader" },
        },
        {
          test: /\.css$/i,
          use: ["style-loader", "css-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/i,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
        {
          test: /\.css$/i,
          use: ["style-loader", "css-loader"],
        },
      ],
    },

    plugins: [
      // ❌ Quitamos CustomFunctionsMetadataPlugin: NO lo usamos.
      // ❌ Quitamos HtmlWebpackPlugin para functions.html.

      // Taskpane
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),

      // Commands (si los usas)
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),

      // Copiamos solo lo necesario a /assets y (opcional) el manifest
      new CopyWebpackPlugin({
        patterns: [
          // 👇 Estas dos son las CLAVE para tus Custom Functions
          { from: path.resolve(__dirname, "assets", "functions.js"), to: "assets/functions.js" },
          { from: path.resolve(__dirname, "assets", "custom-functions.json"), to: "assets/custom-functions.json" },

          // Íconos u otros assets
          { from: path.resolve(__dirname, "assets", "logo-filled.png"), to: "assets/logo-filled.png" },

          // (Opcional) Copiar manifest al output para inspección
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              // En producción reemplaza URLs dev por prod (y quita /public si estuviera)
              return dev
                ? content
                : content.toString().replace(new RegExp(urlDev + "(?:public/)?", "g"), urlProd);
            },
          },
        ],
      }),
    ],

    // Webpack Dev Server v5 con HTTPS (lee certificados desde variables)
    devServer: {
      server: {
        type: "https",
        options: {
          key: fs.readFileSync(process.env.SSL_KEY_FILE),
          cert: fs.readFileSync(process.env.SSL_CRT_FILE),
        },
      },
      host: "localhost", // si tu IT molesta con ::1, cambia a "127.0.0.1"
      port: 3000,
      hot: true,

      // Sirve la RAÍZ del proyecto -> /assets/... existe de verdad
      static: [{ directory: path.join(__dirname), publicPath: "/" }],
    },
  };
};
