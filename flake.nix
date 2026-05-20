{
  description = "A pure Nix flake for a Rust/Slint project";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
    rust-overlay.url = "github:oxalica/rust-overlay";
    flake-utils.url = "github:numtide/flake-utils";
  };

  outputs = { self, nixpkgs, rust-overlay, flake-utils, ... }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        overlays = [ (import rust-overlay) ];
        pkgs = import nixpkgs {
          inherit system overlays;
        };

        # Rust toolchain setup
        rustToolchain = pkgs.rust-bin.stable.latest.default.override {
          extensions = [ "rust-src" "rust-analyzer" "clippy" ];
        };

        # Libraries required at runtime by wayland/opengl/etc.
        runtimeLibs = with pkgs; [
          wayland
          libxkbcommon
          libGL
          fontconfig
          libX11
          libXcursor
          libXi
          libXrandr
        ];

        # Build-time tools and dependencies
        buildInputs = with pkgs; [
          pkg-config
          mold
          clang

          # Rust tools previously in devbox
          bacon
          cargo-release
          cargo-about
          cargo-audit
          cargo-cyclonedx
          cargo-deny
          cargo-edit
          cargo-expand
          cargo-license
          cargo-llvm-cov
          cargo-nextest
          sccache
          prek
          slint-lsp
        ] ++ runtimeLibs;

      in {
        devShells.default = pkgs.mkShell {
          buildInputs = buildInputs;

          nativeBuildInputs = [ rustToolchain ];

          # Environment variables
          CC = "clang";
          CXX = "clang++";
          CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_LINKER = "clang";
          CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS = "-C link-arg=-fuse-ld=mold";
          RUSTC_WRAPPER = "sccache";

          # Ensure runtime libraries can be found by dynamically loaded libraries
          LD_LIBRARY_PATH = "${pkgs.lib.makeLibraryPath runtimeLibs}:/run/opengl-driver/lib:/run/opengl-driver-32/lib";

          # Shell hook to run when entering the shell
          shellHook = ''
            echo '🦀 Rust Dev Workspace | Pure Nix Flake'
            echo 'Available scripts:'
            echo '  dev    - SLINT_LIVE_PREVIEW=1 cargo run -p gui --features dev-ui'
            echo '  test   - cargo nextest run'
            echo '  lint   - cargo clippy'
            echo '  fmt    - cargo fmt'

            # Create aliases to mimic devbox scripts functionality
            alias dev="SLINT_LIVE_PREVIEW=1 cargo run -p gui --features dev-ui"
            alias test="cargo nextest run"
            alias lint="cargo clippy"
            alias fmt="cargo fmt"
          '';
        };
      });
}
