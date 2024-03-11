{
  description = "Script for automatically looking through the CSSE120 gradebook for missing grade items";

  inputs.nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";

  outputs = { self, nixpkgs, ... }:
    let
      system = "x86_64-linux";
      pkgs = import nixpkgs { inherit system; };
    in {
      packages.${system}.default = let
        name = "CSSE120AutoGradebook";
        version = "0.1.0";
        pyproject = (pkgs.formats.toml { }).generate "pyproject.toml" {
          project = {
            inherit name version;
            readme = "README.md";
            # dependencies = [ "tkinter" ];
          };
        };
      in with pkgs.python311Packages; buildPythonPackage {
        inherit name version;
        src = ./.;
        format = "pyproject";
        preBuild = "cp ${pyproject} pyproject.toml";
        buildInputs = [ setuptools tkinter ];
      };

      formatter.${system} = pkgs.nixpkgs-fmt;

      devShells.${system}.default = pkgs.mkShell {
        buildInputs = [
          (pkgs.python3.withPackages (py-pkgs: with py-pkgs; [
            tkinter
            jedi-language-server
          ]))
        ];
      };
    };
}
