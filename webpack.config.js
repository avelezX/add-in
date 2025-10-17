/* eslint-disable no-undef */
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const path = require("path");
const urlProd = "https://avelezx.github.io/add-in/";

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
      path: path.resolve(__dirname, "dist"), // 📁 donde Webpack colocará los archivos compilados
      publicPath: urlProd, // 🌐 base URL de tu add-in en GitHub Pages
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
              return content
              .toString()
              .replace(/https:\/\/localhost:3000\//g, urlProd)
              .replace(/https:\/\/www\.contoso\.com\//g, urlProd);
            },
          },
        ],
      }),
    ],
  };
};
