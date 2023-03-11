"""Command-line interface."""
import click


@click.command()
@click.version_option()
def main() -> None:
    """Kompas 3D Wrapper."""


if __name__ == "__main__":
    main(prog_name="kompas-3d-wrapper")  # pragma: no cover
