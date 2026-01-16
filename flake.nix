{
  description = "trade-analysis flake";

  inputs = {
    nixpkgs.url = "github:nixos/nixpkgs/nixos-unstable";
    flake-parts.url = "github:hercules-ci/flake-parts";
  };

  outputs = { self, nixpkgs, flake-parts, ... }:
    flake-parts.lib.mkFlake {
      inherit self nixpkgs;

      systems =
        [ "x86_64-linux" "x86_64-darwin" "aarch64-linux" "aarch64-darwin" ];

      perSystem = { system, ... }:
        let
          pkgs = import nixpkgs { inherit system; };

          envDeps = with pkgs; [
            pkgs.python3.withPackages
            (p:
              with p; [
                numpy
                pandas
                openpyxl
                beautifulsoup4
                selenium
                thefuzz
                requests
                lxml
                xlsxwriter
              ])
            firefox
            geckodriver
          ];
        in {
          packages.default = pkgs.stdenv.mkDerivation {
            pname = "nvcr trade-analysis";
            version = "0.1";

            src = ./.;

            buildInputs = envDeps;
            dontBuild = true;

            installPhase = ''
              mkdir -p $out/bin
              cp $src/*.py $out/bin/
              mv $out/bin/trade_analysis.py $out/bin/trade-analysis
              chmod +x $out/bin/*
            '';
          };

          devShells.default = pkgs.mkShell { buildInputs = envDeps; };
        };
    };
}
