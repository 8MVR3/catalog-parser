"""Парсер на подменённых страницах: проверяем пагинацию, поля и сохранение в Excel без сети."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import parse_books as P  # noqa: E402

LIST = """<html><body>{cards}<ul class="pager">{nxt}</ul></body></html>"""
CARD = """<article class="product_pod"><h3><a href="{slug}/index.html" title="{title}">Book...</a></h3>
<div class="product_price"><p class="price_color">£{price}</p><p class="instock availability">In stock</p></div>
<p class="star-rating {rating}"></p></article>"""
DETAIL = """<html><body><ul class="breadcrumb"><li><a href="../../index.html">Home</a></li>
<li><a href="../category/books_1/index.html">Books</a></li>
<li><a href="../category/books/poetry_23/index.html">Poetry</a></li><li class="active">{title}</li></ul>
<table class="table table-striped"><tr><th>UPC</th><td>{upc}</td></tr>
<tr><th>Availability</th><td>In stock ({stock} available)</td></tr></table></body></html>"""

PAGES = {}
for p in (1, 2):
    cards = "".join(CARD.format(slug=f"book-{p}-{i}_{p}{i}", title=f"Book {p}-{i}", price=f"{10 + p + i}.50",
                                rating=["One", "Three", "Five"][i]) for i in range(3))
    nxt = '<li class="next"><a href="page-2.html">next</a></li>' if p == 1 else ""
    PAGES[f"https://books.toscrape.com/catalogue/page-{p}.html"] = LIST.format(cards=cards, nxt=nxt)


class FakeResp:
    def __init__(self, text):
        self.text, self.encoding = text, None

    def raise_for_status(self):
        pass


class FakeSession:
    headers = {}

    def get(self, url, timeout=20):
        if url in PAGES:
            return FakeResp(PAGES[url])
        return FakeResp(DETAIL.format(title="x", upc=f"u{len(url)}{url[-12:]}", stock=7))


def test_parse_and_save(tmp_path, monkeypatch):
    monkeypatch.setattr(P.time, "sleep", lambda *_: None)
    session = FakeSession()
    books = P.parse_list(session)
    assert len(books) == 6
    P.parse_details(session, books)
    assert all(b["category"] == "Poetry" and b["stock"] == 7 and b["upc"] for b in books)
    assert books[0]["price"] == 11.5 and books[0]["rating"] == 1
    out = tmp_path / "books.xlsx"
    P.save(books, True, str(out))
    assert out.exists() and out.stat().st_size > 0
