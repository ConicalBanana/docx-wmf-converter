import os
import typing as t
from functools import wraps
import tempfile
import zipfile
import shutil
from pathlib import Path

from bs4 import BeautifulSoup


def with_temp_path(func: t.Callable) -> t.Callable:
    @wraps(func)
    def wrapper(*args, **kwargs):
        if kwargs.get("temp_path") is None:
            kwargs["temp_path"] = Path(tempfile.mkdtemp())
        try:
            return func(*args, **kwargs)
        finally:
            shutil.rmtree(kwargs["temp_path"])
    return wrapper


@with_temp_path
def replace_fig(
    inp: str,
    out: str,
    **kwargs,
):
    '''Replace all the EMF/WMF figures in the docx file with SVG figures.

    Args:
        inp: The input path.
        out: The output path.
    '''
    # From decorator
    tmp_path: Path = kwargs["temp_path"]

    # Unzip the input file
    with zipfile.ZipFile(inp, "r") as fo:
        fo.extractall(tmp_path)

    # Image references record.
    doc_rel_path = tmp_path / "word" / "_rels" / "document.xml.rels"
    doc_content_type_path = tmp_path / "[Content_Types].xml"
    img_path = tmp_path / "word" / "media"
    img_list = [i for i in img_path.iterdir()]

    # Read document layout
    with open(doc_rel_path, "r") as fo:
        doc_rel_content = fo.read()

    # Modify resourse
    for img in img_list:
        print(img)
        if img.suffix == ".emf":
            img_from = img.name
            img_to = img.name.replace(".emf", ".svg")
            img_to_path = img.parent / img_to
            print("Replacing image", img_from, "with", img_to)

            os.system(f"emf2svg-conv -i {img} -o {img_to_path}")
            img.unlink()

            doc_rel_content = doc_rel_content.replace(img_from, img_to)
        elif img.suffix == ".wmf":
            img_from = img.name
            img_to = img.name.replace(".wmf", ".svg")
            img_to_path = img.parent / img_to
            print("Replacing image", img_from, "with", img_to)

            os.system(f"wmf2svg {img} -o {img_to_path}")
            img.unlink()

            doc_rel_content = doc_rel_content.replace(img_from, img_to)

    # Save layout modifications
    doc_rel_path.unlink()
    with open(doc_rel_path, "w") as fo:
        fo.write(doc_rel_content)

    # Add svg extension
    with open(doc_content_type_path, "r") as fo:
        soup = BeautifulSoup(fo.read(), "xml")
    root = soup.find("Types")
    svg_tag = soup.new_tag(
        "Default",
        attrs={
            "Extension": "svg",
            "ContentType": "image/svg+xml"
        },
    )
    root.find("Default").insert_after(svg_tag)
    doc_content_type_path.unlink()
    with open(doc_content_type_path, "w") as fo:
        fo.write(soup.prettify())

    # Repack
    current_dir = os.getcwd()
    with zipfile.ZipFile(out, "w") as fo:
        os.chdir(tmp_path)
        for dir_path, _, file_names in tmp_path.walk():
            for file_name in file_names:
                item = dir_path / file_name
                print(item.relative_to(tmp_path))
                fo.write(item.relative_to(tmp_path))
    os.chdir(current_dir)


if __name__ == "__main__":
    replace_fig(
        inp="test/test.docx",
        out="test/test_output.docx",
    )
