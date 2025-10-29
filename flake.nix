{
  description = "trade-analysis flake";

  inputs = {
    nixpkgs.url = "github:nixos/nixpkgs/nixos-unstable";
  };

  outputs = { nixpkgs, ... }:
    let
    supportedSystems = [ "x86_64-linux" "x86_64-darwin" "aarch64-linux" "aarch64-darwin" ];

  forAllSystems = f: 
    nixpkgs.lib.genAttrs supportedSystems ( system: 
	let
	  pkgs = nixpkgs.legacyPackages.${system};

	  pythonEnv = pkgs.python3.withPackages ( p: with p; [
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

	in 

	  f { inherit pkgs pythonEnv system; }
    );

  in {
    packages = forAllSystems ( { pkgs, pythonEnv, ... }: { 
	default = pkgs.stdenv.mkDerivation {
	  pname = "nvcr trade-analysis";
	  version = "0.1";

	  src = ./.;

	  buildInputs = [ pythonEnv pkgs.geckodriver pkgs.firefox];
	  dontBuild = true;

	  installPhase = ''
	    mkdir -p $out/bin
	    cp $src/db-nvrmap.py $out/bin/db-nvrmap
	    cp $src/db-ensym.py $out/bin/db-ensym
	    chmod +x $out/bin/db-nvrmap
	    chmod +x $out/bin/db-ensym
	  '';
	};
    });

    devShells = forAllSystems ({ pkgs, pythonEnv, ... }: {
	default = pkgs.mkShell {
	  buildInputs = [ pythonEnv pkgs.geckodriver pkgs.firefox ];
	};
    });
  };
}
