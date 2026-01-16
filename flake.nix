{
  description = "trade-analysis flake";

  inputs = { nixpkgs.url = "github:nixos/nixpkgs/nixos-unstable"; };

  outputs = { nixpkgs, ... }:
    let
      supportedSystems =
        [ "x86_64-linux" "x86_64-darwin" "aarch64-linux" "aarch64-darwin" ];

      forAllSystems = f:
        nixpkgs.lib.genAttrs supportedSystems (system:
          let
            pkgs = nixpkgs.legacyPackages.${system};

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

            envDeps = with pkgs; [ pythonEnv firefox geckodriver ];

          in f { inherit pkgs envDeps system; });

    in {
      packages = forAllSystems ({ pkgs, envDeps, ... }: {
        default = pkgs.stdenv.mkDerivation {
          pname = "nvcr trade-analysis";
          version = "0.1";

          src = ./.;

          buildInputs = envDeps;
          dontBuild = true;

          installPhase =
            "	mkdir -p $out/bin\n	cp $src/*.py $out/bin/\n	mv $out/bin/trade_analysis.py $out/bin/trade_analysis\n	chmod +x $out/bin/*.py\n	chmod +x $out/bin/trade_analysis\n";
        };
      });

      devShells = forAllSystems ({ pkgs, envDeps, ... }: {
        default = pkgs.mkShell { buildInputs = [ envDeps ]; };
      });
    };
}
