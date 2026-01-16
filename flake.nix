{
  description = "trade-analysis flake";

  inputs = {
    nixpkgs.url = "github:nixos/nixpkgs/nixos-unstable";
    flake-parts.url = "github:hercules-ci/flake-parts";
  };

  outputs = inputs@{ self, nixpkgs, flake-parts, ... }:
    flake-parts.lib.mkFlake { inherit inputs; } {
      systems =
        [ "x86_64-linux" "x86_64-darwin" "aarch64-linux" "aarch64-darwin" ];

      perSystem = { pkgs, system, lib, ... }:
        let
          pythonEnv = pkgs.python3.withPackages (p:
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
            ]);

          runtimeDeps = [ pkgs.firefox pkgs.geckodriver ];
        in {
          packages.default = pkgs.stdenvNoCC.mkDerivation {
            pname = "nvcr-trade-analysis";
            version = "0.1";

            src = ./.;

            nativeBuildInputs = [ pkgs.makeWrapper ];
            dontBuild = true;

            installPhase = ''
              runHook preInstall
              mkdir -p "$out/bin" "$out/lib/trade-analysis"
              cp "$src"/*.py "$out/lib/trade-analysis/"

              makeWrapper "${pythonEnv}/bin/python" "$out/bin/trade-analysis" \
                --add-flags "$out/lib/trade-analysis/trade_analysis.py" \
                --prefix PATH : ${lib.makeBinPath runtimeDeps}

              chmod +x "$out/bin/trade-analysis"
              runHook postInstall
            '';
          };

          devShells.default =
            pkgs.mkShell { packages = [ pythonEnv ] ++ runtimeDeps; };
        };
    };
}
