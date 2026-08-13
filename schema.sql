CREATE TABLE IF NOT EXISTS books_toscrape (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT,
    price REAL,
    availability TEXT,
    star_rating TEXT,
    book_url TEXT,
    page INTEGER,
    source TEXT
);

CREATE TABLE IF NOT EXISTS gaaraas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT,
    price TEXT,
    location TEXT,
    details TEXT,
    listing_url TEXT,
    page INTEGER,
    source TEXT
);
