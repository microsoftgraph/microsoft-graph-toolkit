import resolve from "rollup-plugin-node-resolve";
import commonJS from "rollup-plugin-commonjs";
import typescript from "rollup-plugin-typescript";
import postcss from "rollup-plugin-postcss";
import { terser } from "rollup-plugin-terser";

const core_opts = {
  plugins: [
    resolve({
      module: true,
      jsnext: true
    }),
    commonJS(),
    terser(),
    typescript({
      allowSyntheticDefaultImports: true,
      experimentalDecorators: true,
      lib: ["dom", "es2015"],
      module: "es2015",
      target: "es2015",
      moduleResolution: "node",
      declaration: true,
      esModuleInterop: true,
      sourceMap: true,
      rootDir: "src"
    }),
    postcss({
      inject: false
    })
  ]
};

export default [
  {
    ...core_opts,
    output: {
      dir: "./bin/components",
      entryFileNames: "[name].js",
      format: "esm"
    },
    input: [
      "./src_redux/components/providers/mgt-mock-provider.ts",
      "./src_redux/components/providers/mgt-msal-provider.ts",
      "./src_redux/components/providers/mgt-teams-provider.ts",
      "./src_redux/components/providers/mgt-wam-providers.ts",

      "./src_redux/components/ui/mgt-agenda/mgt-agenda.ts",
      "./src_redux/components/ui/mgt-person/mgt-person.ts",
      "./src_redux/components/ui/mgt-login/mgt-login.ts"
    ]
  },
  {
    ...core_opts,
    output: {
      dir: "./bin/bundle",
      entryFileNames: "[name].js",
      format: "esm"
    },
    input: "./src_redux/index.ts"
  }
];
