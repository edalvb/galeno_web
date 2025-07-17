// tailwind.config.js
const {heroui} = require("@heroui/react");

/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    // ...
    // make sure it's pointing to the ROOT node_module
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
    "./node_modules/@heroui/theme/dist/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {},
    fontFamily: {
      sans: ['Sora', 'sans-serif'],
    },
  },
  darkMode: "class",
  plugins: [heroui({
    themes: {
      primaryMainTheme : {
        extend : "light",
        colors: {
          background: "#ffffff",
          foreground: "#11181c",
          primary : {
            DEFAULT: "#BD273C",
            foreground: "#ffffff",
          },
          secondary : {
            DEFAULT: "#0C1421",
            foreground: "#ffffff",
          },
          third : {
            DEFAULT: "#8897AD",
            foreground: "#ffffff",
          },
          hover_highlight: "#a1a1aa33",
          background_secondary: "#F5F6F9",
          middle_ground: "#5c5d67",
          button_secondary: "#D8D9DA",
          light_box: "#E6E6F1",
          gray: {
            DEFAULT: "#7D7E8A",
          }
        }
      },
      twLight : {
        extend : "light",
        colors: {
          background: "#ffffff",
          foreground: "#1C1A1C",
          primary : {
            DEFAULT: "#000f96",
            foreground: "#ffffff",
          },
          background_secondary: "#8ab2e8",
          middle_ground: "#7D7E8A",
          button_secondary: "#D8D9DA",
        },
      },
      twDark : {
        extend : "dark",
        colors: {
          background: "#11082d",
          foreground: "#f84bd5",
          primary : {
            DEFAULT: "#1b26bc",
            foreground: "#ffffff",
          },
          background_secondary: "#0c0233",
        }
      }
    }
  })],
};