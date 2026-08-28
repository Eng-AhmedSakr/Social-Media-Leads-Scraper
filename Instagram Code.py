import asyncio
import re
import pandas as pd
from pathlib import Path
from playwright.async_api import async_playwright


# ============================================================
# SETTINGS
# ============================================================

EXCEL_FILE = "input_data.xlsx"
OUTPUT_FILE = "output_results.xlsx"

# عدد الشركات
MAX_COMPANIES = 76

# عدد البوستات من كل صفحة
MAX_POSTS_PER_PAGE = 20

# المتصفح يظهر أمامك
HEADLESS = False

# بروفايل Instagram المحفوظ
PROFILE_DIR = Path("./browser_profile")

# الانتظار بين العمليات
WAIT_AFTER_PROFILE = 3
WAIT_AFTER_POST = 4
WAIT_AFTER_CLICK = 1

# عدد مرات محاولة تحميل الكومنتات
COMMENT_LOAD_ROUNDS = 8

# عدد مرات Scroll داخل البوست
COMMENT_SCROLL_ROUNDS = 15


# ============================================================
# IGNORED USERNAMES
# ============================================================

IGNORED_USERNAMES = {
    "instagram",
    "meta",
    "about",
    "blog",
    "jobs",
    "help",
    "privacy",
    "terms",
    "locations",
    "popular",
    "explore",
    "direct",
    "accounts",
    "settings",
    "challenge",
    "login",
    "reels",
    "stories",
    "web",
}


# ============================================================
# IGNORED TEXT
# ============================================================

IGNORED_TEXT = {
    "Log In",
    "Sign Up",
    "Follow",
    "Following",
    "Message",
    "Like",
    "Likes",
    "Reply",
    "Replies",
    "Share",
    "Save",
    "Send",
    "Post",
    "Translate",
    "See translation",
    "Edited",
    "More",
    "View more",
    "View replies",
    "View more replies",
    "Load more replies",
    "View all comments",
    "View more comments",
    "Load more comments",
    "Hide replies",
    "Write a comment...",
    "Add a comment...",
    "No comments yet",
    "No comments yet.",
    "Original audio",
    "Original sound",
    "Paid partnership with",
    "More posts from",
    "Meta",
    "About",
    "Blog",
    "Jobs",
    "Help",
    "API",
    "Privacy",
    "Terms",
    "Locations",
    "Popular",
    "Instagram Lite",
    "Meta AI",
    "Threads",
    "Contact Uploading & Non-Users",
    "Meta Verified",
    "Messages",
    "Start the conversation.",
    "AI content",
    "English",
    "Afrikaans",
    "العربية",
    "Čeština",
    "Dansk",
    "Deutsch",
    "Ελληνικά",
    "Español",
    "فارسی",
    "Français",
    "עברית",
    "Italiano",
    "日本語",
    "한국어",
    "Português",
    "Русский",
    "Türkçe",
    "中文(简体)",
}


# ============================================================
# TIME LABELS
# ============================================================

TIME_LABELS = set()

for i in range(1, 61):
    TIME_LABELS.add(f"{i}m")

for i in range(1, 25):
    TIME_LABELS.add(f"{i}h")

for i in range(1, 32):
    TIME_LABELS.add(f"{i}d")

for i in range(1, 53):
    TIME_LABELS.add(f"{i}w")

for i in range(1, 25):
    TIME_LABELS.add(f"{i}mo")

for i in range(1, 6):
    TIME_LABELS.add(f"{i}y")

TIME_LABELS.add("now")

TIME_LABELS_LOWER = {
    x.lower()
    for x in TIME_LABELS
}


# ============================================================
# TEXT HELPERS
# ============================================================

def normalize_text(text):

    if text is None:
        return ""

    text = str(text)

    text = text.replace("\u200b", " ")
    text = text.replace("\xa0", " ")

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# DATE CHECK
# ============================================================

def is_date(text):

    text = normalize_text(text)

    patterns = [

        r"^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}$",

        r"^\d{1,2}/\d{1,2}/\d{4}$",

        r"^\d{1,2}-\d{1,2}-\d{4}$",

        r"^\d{4}-\d{1,2}-\d{1,2}$",

        r"^\d+\s+(day|days|week|weeks|month|months|year|years)\s+ago$",

    ]

    for pattern in patterns:

        if re.match(
            pattern,
            text,
            re.I
        ):
            return True

    return False


# ============================================================
# TIME CHECK
# ============================================================

def is_time_label(text):

    text = normalize_text(text)

    return (
        text.lower()
        in TIME_LABELS_LOWER
    )


# ============================================================
# LANGUAGE CHECK
# ============================================================

def contains_arabic(text):

    return bool(
        re.search(
            r"[\u0600-\u06FF]",
            text
        )
    )


def contains_english(text):

    return bool(
        re.search(
            r"[A-Za-z]",
            text
        )
    )


# ============================================================
# OBVIOUS UI
# ============================================================

def is_obvious_ui(text):

    text = normalize_text(text)

    if not text:
        return True

    low = text.lower()

    ignored = {
        x.lower()
        for x in IGNORED_TEXT
    }

    if low in ignored:
        return True

    if is_time_label(text):
        return True

    if is_date(text):
        return True

    # Instagram counters
    if re.match(
        r"^\d[\d,.\s]*\s+(likes?|comments?|replies?)$",
        text,
        re.I
    ):
        return True

    # Reply counters
    if re.match(
        r"^(view\s+(all\s+)?\d+\s+repl(?:y|ies)|\d+\s+repl(?:y|ies))$",
        text,
        re.I
    ):
        return True

    # "and 123 others"
    if re.match(
        r"^and\s+\d+\s+others?$",
        text,
        re.I
    ):
        return True

    # Pure number / punctuation
    if re.fullmatch(
        r"[\d\s.,:/\-]+",
        text
    ):
        return True

    # UI phrases
    bad_phrases = [

        "more posts from",
        
        
        
        
        "instagram from meta",
        "contact uploading",
        "terms of use",
        "privacy policy",
        "copyright",
        "© instagram",
        "original audio",
        "original sound",
        "paid partnership",

    ]

    for phrase in bad_phrases:

        if phrase in low:
            return True

    return False


# ============================================================
# COMMENT CHECK
# ============================================================

def looks_like_comment(text):

    text = normalize_text(text)

    if not text:
        return False

    if is_obvious_ui(text):
        return False

    if len(text) < 2:
        return False

    if len(text) > 1000:
        return False

    # pure @username
    if re.fullmatch(
        r"@[\w.\-]+",
        text
    ):
        return False

    # URL
    if re.fullmatch(
        r"https?://\S+",
        text,
        re.I
    ):
        return False

    # Pure username
    if re.fullmatch(
        r"[A-Za-z0-9._]+",
        text
    ):
        return False

    # Arabic
    if contains_arabic(text):
        return True

    # English
    if contains_english(text):

        rejected = {
            "and",
            "edited",
            "more",
            "follow",
            "following",
            "like",
            "likes",
            "share",
            "reply",
            "replies",
            "send",
            "post",
            "please",
        }

        if text.lower() in rejected:
            return False

        return True

    # Emoji / symbols only
    if re.fullmatch(
        r"[\W_]+",
        text,
        re.UNICODE
    ):
        return True

    return False


# ============================================================
# SAFE WAIT
# ============================================================

async def safe_wait(
    page,
    milliseconds=2000
):

    try:

        await page.wait_for_timeout(
            milliseconds
        )

    except Exception:

        pass


# ============================================================
# LOAD EXCEL
# ============================================================

def load_companies():

    print("=" * 70)
    print("READING EXCEL")
    print("=" * 70)

    path = Path(
        EXCEL_FILE
    )

    if not path.exists():

        raise FileNotFoundError(
            f"\nExcel file not found:\n{EXCEL_FILE}"
        )

    raw = pd.read_excel(
        EXCEL_FILE,
        header=None
    )

    header_row = None

    # البحث عن Header
    for i in range(
        len(raw)
    ):

        row = [
            normalize_text(x)
            for x in raw.iloc[i].tolist()
        ]

        if (
            "Company Name" in row
            and
            "Instagram" in row
        ):

            header_row = i
            break

    if header_row is None:

        raise ValueError(
            """
Could not find Excel header row.

Expected columns:

Company Name
Instagram
"""
        )

    print(
        f"\nHeader row found at: {header_row}"
    )

    df = pd.read_excel(
        EXCEL_FILE,
        header=header_row
    )

    # تنظيف أسماء الأعمدة
    df.columns = [
        normalize_text(col)
        for col in df.columns
    ]

    print(
        "\nExcel columns:",
        list(df.columns)
    )

    required = [
        "Company Name",
        "Instagram"
    ]

    for column in required:

        if column not in df.columns:

            raise ValueError(
                f"""
Missing column: {column}

Available columns:
{list(df.columns)}
"""
            )

    df = df[
        required
    ].copy()

    # تنظيف البيانات
    df["Company Name"] = (
        df["Company Name"]
        .fillna("")
        .astype(str)
        .map(normalize_text)
    )

    df["Instagram"] = (
        df["Instagram"]
        .fillna("")
        .astype(str)
        .map(normalize_text)
    )

    # إزالة الصفوف الفارغة
    df = df[
        df["Company Name"].str.strip() != ""
    ]

    df = df[
        df["Instagram"].str.contains(
            "instagram.com",
            case=False,
            na=False
        )
    ]

    # إزالة التكرار
    df = df.drop_duplicates(
        subset=[
            "Instagram"
        ]
    )

    # العدد المطلوب
    df = df.head(
        MAX_COMPANIES
    )

    print(
        f"\nTotal companies: {len(df)}"
    )

    print(
        f"Maximum posts per page: {MAX_POSTS_PER_PAGE}"
    )

    print(
        f"Theoretical maximum posts: "
        f"{len(df) * MAX_POSTS_PER_PAGE}"
    )

    return df.reset_index(
        drop=True
    )


# ============================================================
# OPEN PROFILE
# ============================================================

async def open_profile(
    context,
    url,
    company
):

    print()
    print("=" * 70)
    print("OPENING PROFILE")
    print("=" * 70)

    print(
        f"Company: {company}"
    )

    print(
        f"URL: {url}"
    )

    page = await context.new_page()

    try:

        await page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=60000
        )

    except Exception as e:

        print(
            "Profile navigation warning:",
            e
        )

    await safe_wait(
        page,
        5000
    )

    print(
        "Current URL:",
        page.url
    )

    return page


# ============================================================
# GET POST LINKS
# ============================================================

async def get_post_links(
    page
):

    print()
    print("Getting post links...")

    previous_count = 0
    stable_rounds = 0

    # Scroll الصفحة
    for round_number in range(40):

        try:

            await page.mouse.wheel(
                0,
                3000
            )

        except Exception:
            pass

        await safe_wait(
            page,
            1200
        )

        try:

            count = await page.locator(
                'a[href*="/p/"], a[href*="/reel/"]'
            ).count()

        except Exception:

            count = previous_count

        print(
            f"  Scroll {round_number + 1}: "
            f"{count} post links"
        )

        if count == previous_count:

            stable_rounds += 1

        else:

            stable_rounds = 0

        previous_count = count

        if count >= MAX_POSTS_PER_PAGE:
            break

        if stable_rounds >= 7:
            break

    # استخراج الروابط
    try:

        links = await page.locator(
            "a[href]"
        ).evaluate_all(
            """
            elements => {

                return elements
                    .map(a => a.href)
                    .filter(href =>
                        href &&
                        (
                            href.includes('/p/') ||
                            href.includes('/reel/')
                        )
                    );

            }
            """
        )

    except Exception as e:

        print(
            "Error getting post links:",
            e
        )

        return []

    unique = []
    seen = set()

    for link in links:

        clean = (
            str(link)
            .split("?")[0]
            .rstrip("/")
            + "/"
        )

        if clean not in seen:

            seen.add(clean)

            unique.append(
                clean
            )

    unique = unique[
        :MAX_POSTS_PER_PAGE
    ]

    print()
    print(
        f"Posts selected: {len(unique)}"
    )

    return unique


# ============================================================
# OPEN POST
# ============================================================

async def open_post(
    context,
    url
):

    page = await context.new_page()

    try:

        await page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=60000
        )

    except Exception as e:

        print(
            "Post navigation warning:",
            e
        )

    await safe_wait(
        page,
        WAIT_AFTER_POST * 1000
    )

    return page


# ============================================================
# CLICK COMMENT BUTTONS
# ============================================================

async def click_comment_buttons(
    page
):

    buttons = [

        "View all comments",
        "View more comments",
        "Load more comments",
        "View more",
        "more comments",
        "View replies",
        "View more replies",
        "Load more replies",
        "Show more",

    ]

    for round_number in range(
        COMMENT_LOAD_ROUNDS
    ):

        clicked = 0

        for text in buttons:

            try:

                locator = page.get_by_text(
                    text,
                    exact=False
                )

                count = await locator.count()

                for i in range(
                    min(count, 20)
                ):

                    try:

                        element = locator.nth(i)

                        if await element.is_visible():

                            await element.scroll_into_view_if_needed(
                                timeout=2000
                            )

                            await element.click(
                                timeout=3000
                            )

                            clicked += 1

                            await safe_wait(
                                page,
                                WAIT_AFTER_CLICK * 1000
                            )

                    except Exception:
                        continue

            except Exception:
                continue

        print(
            f"  Comment load round "
            f"{round_number + 1}: "
            f"{clicked} clicks"
        )

        # Scroll
        try:

            await page.mouse.wheel(
                0,
                1200
            )

        except Exception:
            pass

        await safe_wait(
            page,
            900
        )


# ============================================================
# EXTRACT USERNAME FROM HREF
# ============================================================

def username_from_href(
    href
):

    if not href:
        return None

    href = str(href).strip()

    # Absolute Instagram URL
    match = re.search(
        r"instagram\.com/([A-Za-z0-9._]{1,30})/?(?:\?.*)?$",
        href,
        re.I
    )

    if match:

        username = match.group(1)

    else:

        # Relative
        if not href.startswith("/"):
            return None

        path = href.split("?")[0]

        parts = [
            x
            for x in path.strip("/").split("/")
            if x
        ]

        if len(parts) != 1:
            return None

        username = parts[0]

    if not username:
        return None

    if username.lower() in IGNORED_USERNAMES:
        return None

    if not re.fullmatch(
        r"[A-Za-z0-9._]{1,30}",
        username
    ):
        return None

    return username


# ============================================================
# FIND PROFILE LINKS
# ============================================================

async def get_profile_links(
    page
):

    try:

        links = await page.locator(
            "a[href]"
        ).evaluate_all(
            """
            elements => {

                return elements.map(a => ({
                    href: a.getAttribute('href') || '',
                    text: a.innerText || ''
                }));

            }
            """
        )

    except Exception:

        return []

    results = []

    seen = set()

    for item in links:

        href = item.get(
            "href",
            ""
        )

        username = username_from_href(
            href
        )

        if not username:
            continue

        profile_url = (
            "https://www.instagram.com/"
            + username
            + "/"
        )

        key = profile_url.lower()

        if key in seen:
            continue

        seen.add(key)

        results.append(
            {
                "username": username,
                "profile_url": profile_url,
                "href": href,
            }
        )

    return results


# ============================================================
# FIND COMMENT FROM CONTAINER
# ============================================================

async def extract_comment_from_container(
    element
):

    try:

        text = await element.inner_text(
            timeout=2000
        )

    except Exception:

        return None

    text = normalize_text(
        text
    )

    if not text:
        return None

    lines = [
        normalize_text(x)
        for x in text.splitlines()
        if normalize_text(x)
    ]

    # Remove duplicate lines
    clean_lines = []

    seen = set()

    for line in lines:

        key = line.lower()

        if key in seen:
            continue

        seen.add(key)

        clean_lines.append(
            line
        )

    # Try every line
    for line in clean_lines:

        if not line:
            continue

        if is_obvious_ui(line):
            continue

        if is_time_label(line):
            continue

        if is_date(line):
            continue

        # Pure username
        if re.fullmatch(
            r"@?[A-Za-z0-9._]+",
            line
        ):
            continue

        # Reply buttons
        if re.match(
            r"^(view|load|show|hide).*(reply|repl|comment)",
            line,
            re.I
        ):
            continue

        if looks_like_comment(line):

            return line

    return None


# ============================================================
# FIND COMMENT CONTAINER AROUND PROFILE LINK
# ============================================================

async def find_comment_for_profile_link(
    link
):

    current = link

    # Walk up DOM
    for level in range(1, 12):

        try:

            current = current.locator(
                "xpath=.."
            )

            text = await current.inner_text(
                timeout=1500
            )

            text = normalize_text(
                text
            )

            if not text:
                continue

            # Limit giant page containers
            if len(text) > 3000:
                continue

            comment = await extract_comment_from_container(
                current
            )

            if comment:

                return comment

        except Exception:

            continue

    return None


# ============================================================
# EXTRACT COMMENTS + PROFILES
# ============================================================

async def extract_comments_and_profiles(
    page,
    post_url,
    company
):

    print()
    print(
        "Loading comments..."
    )

    await click_comment_buttons(
        page
    )

    # Additional scrolling
    print(
        "Scrolling comments..."
    )

    for i in range(
        COMMENT_SCROLL_ROUNDS
    ):

        try:

            await page.mouse.wheel(
                0,
                1400
            )

        except Exception:
            pass

        await safe_wait(
            page,
            700
        )

    # --------------------------------------------------------
    # Get profile links
    # --------------------------------------------------------

    profiles = await get_profile_links(
        page
    )

    print(
        f"Profile links found on page: "
        f"{len(profiles)}"
    )

    results = []

    seen = set()

    profiles_found = 0
    comments_found = 0

    # --------------------------------------------------------
    # Process profile links
    # --------------------------------------------------------

    for profile in profiles:

        username = profile[
            "username"
        ]

        profile_url = profile[
            "profile_url"
        ]

        href = profile[
            "href"
        ]

        # Find actual DOM link
        try:

            locator = page.locator(
                f'a[href="{href}"]'
            )

            count = await locator.count()

        except Exception:

            count = 0

        if count == 0:

            # Try generic href contains
            try:

                locator = page.locator(
                    f'a[href*="{username}"]'
                )

                count = await locator.count()

            except Exception:

                count = 0

        if count == 0:
            continue

        for i in range(
            min(count, 10)
        ):

            try:

                link_element = locator.nth(i)

                if not await link_element.is_visible():
                    continue

                comment = (
                    await find_comment_for_profile_link(
                        link_element
                    )
                )

                if not comment:
                    continue

                # Avoid capturing post owner/caption
                # when container is too broad
                key = (
                    company.lower(),
                    post_url.lower(),
                    profile_url.lower(),
                    comment
                )

                if key in seen:
                    continue

                seen.add(key)

                results.append(
                    {
                        "Page Name": company,
                        "Post URL": post_url,
                        "Comment": comment,
                        "Profile URL": profile_url,
                    }
                )

                profiles_found += 1
                comments_found += 1

                print()
                print(
                    f"COMMENT {comments_found}:"
                )

                print(
                    f"  User: {username}"
                )

                print(
                    f"  Profile: {profile_url}"
                )

                print(
                    f"  Comment: {comment}"
                )

                break

            except Exception:
                continue

    # ========================================================
    # SECOND EXTRACTION METHOD
    # Find elements containing Instagram profile href
    # and inspect nearby DOM
    # ========================================================

    if not results:

        print()
        print(
            "Primary extraction found 0."
        )

        print(
            "Trying secondary DOM extraction..."
        )

        try:

            candidates = await page.locator(
                'a[href^="/"]'
            ).count()

        except Exception:

            candidates = 0

        print(
            f"Candidate profile links: {candidates}"
        )

        for i in range(
            min(candidates, 500)
        ):

            try:

                link = page.locator(
                    'a[href^="/"]'
                ).nth(i)

                href = await link.get_attribute(
                    "href"
                )

                username = username_from_href(
                    href
                )

                if not username:
                    continue

                if not await link.is_visible():
                    continue

                profile_url = (
                    "https://www.instagram.com/"
                    + username
                    + "/"
                )

                comment = (
                    await find_comment_for_profile_link(
                        link
                    )
                )

                if not comment:
                    continue

                key = (
                    company.lower(),
                    post_url.lower(),
                    profile_url.lower(),
                    comment
                )

                if key in seen:
                    continue

                seen.add(key)

                results.append(
                    {
                        "Page Name": company,
                        "Post URL": post_url,
                        "Comment": comment,
                        "Profile URL": profile_url,
                    }
                )

                print()
                print(
                    "SECONDARY COMMENT:"
                )

                print(
                    f"  User: {username}"
                )

                print(
                    f"  Profile: {profile_url}"
                )

                print(
                    f"  Comment: {comment}"
                )

            except Exception:
                continue

    # ========================================================
    # THIRD METHOD
    # Search DOM elements directly for comment-like text
    # ========================================================

    if not results:

        print()
        print(
            "Profile-based extraction found 0."
        )

        print(
            "Trying text/DOM fallback..."
        )

        try:

            elements = await page.locator(
                "div"
            ).evaluate_all(
                """
                elements => {

                    const output = [];

                    for (const el of elements) {

                        const text = (el.innerText || '').trim();

                        if (!text)
                            continue;

                        if (text.length < 2)
                            continue;

                        if (text.length > 1000)
                            continue;

                        const links = Array.from(
                            el.querySelectorAll('a[href]')
                        ).map(a => ({
                            href: a.getAttribute('href') || '',
                            text: a.innerText || ''
                        }));

                        if (links.length === 0)
                            continue;

                        output.push({
                            text: text,
                            links: links
                        });

                    }

                    return output;

                }
                """
            )

        except Exception as e:

            print(
                "DOM fallback error:",
                e
            )

            elements = []

        print(
            f"Potential DOM containers: "
            f"{len(elements)}"
        )

        for element in elements:

            text = normalize_text(
                element.get(
                    "text",
                    ""
                )
            )

            if not text:
                continue

            if len(text) > 1000:
                continue

            links = element.get(
                "links",
                []
            )

            username = None
            profile_url = None

            # Find profile link
            for link_data in links:

                href = link_data.get(
                    "href",
                    ""
                )

                candidate = username_from_href(
                    href
                )

                if candidate:

                    username = candidate

                    profile_url = (
                        "https://www.instagram.com/"
                        + candidate
                        + "/"
                    )

                    break

            if not username:
                continue

            # Find comment line
            lines = [
                normalize_text(x)
                for x in text.splitlines()
                if normalize_text(x)
            ]

            for line in lines:

                if (
                    line.lower()
                    == username.lower()
                ):
                    continue

                if not looks_like_comment(
                    line
                ):
                    continue

                key = (
                    company.lower(),
                    post_url.lower(),
                    profile_url.lower(),
                    line
                )

                if key in seen:
                    continue

                seen.add(key)

                results.append(
                    {
                        "Page Name": company,
                        "Post URL": post_url,
                        "Comment": line,
                        "Profile URL": profile_url,
                    }
                )

                print()
                print(
                    "FALLBACK COMMENT:"
                )

                print(
                    f"  User: {username}"
                )

                print(
                    f"  Profile: {profile_url}"
                )

                print(
                    f"  Comment: {line}"
                )

                break

    # ========================================================
    # FINAL CLEANUP
    # ========================================================

    final_results = []

    final_seen = set()

    for item in results:

        comment = normalize_text(
            item["Comment"]
        )

        profile_url = normalize_text(
            item["Profile URL"]
        )

        post = normalize_text(
            item["Post URL"]
        )

        company_name = normalize_text(
            item["Page Name"]
        )

        if not comment:
            continue

        if not post:
            continue

        if not company_name:
            continue

        if is_obvious_ui(comment):
            continue

        if is_time_label(comment):
            continue

        if is_date(comment):
            continue

        if len(comment) < 2:
            continue

        if len(comment) > 1000:
            continue

        # Remove language selector
        if comment.strip() in {
            "اردو",
            "English",
            "العربية",
            "Deutsch",
            "Français",
            "Español",
            "Italiano",
        }:
            continue

        key = (
            company_name.lower(),
            post.lower(),
            profile_url.lower(),
            comment.lower()
        )

        if key in final_seen:
            continue

        final_seen.add(key)

        final_results.append(
            {
                "Page Name": company_name,
                "Post URL": post,
                "Comment": comment,
                "Profile URL": profile_url,
            }
        )

    print()
    print(
        "=" * 70
    )

    print(
        f"REAL COMMENTS FOUND: "
        f"{len(final_results)}"
    )

    print(
        f"Profiles found: "
        f"{sum(1 for x in final_results if x['Profile URL'])}"
    )

    print(
        f"Profiles not found: "
        f"{sum(1 for x in final_results if not x['Profile URL'])}"
    )

    print(
        "=" * 70
    )

    return final_results


# ============================================================
# SAVE RESULTS
# ============================================================

def save_results(
    results
):

    columns = [
        "Page Name",
        "Post URL",
        "Comment",
        "Profile URL",
    ]

    if results:

        df = pd.DataFrame(
            results
        )

        for column in columns:

            if column not in df.columns:

                df[column] = ""

        df = df[
            columns
        ]

        # Normalize
        for column in columns:

            df[column] = (
                df[column]
                .fillna("")
                .astype(str)
                .map(normalize_text)
            )

        # Remove UI
        df = df[
            ~df["Comment"].apply(
                is_obvious_ui
            )
        ]

        # Remove duplicates
        df = df.drop_duplicates(
            subset=[
                "Page Name",
                "Post URL",
                "Comment",
                "Profile URL",
            ]
        )

    else:

        df = pd.DataFrame(
            columns=columns
        )

    df.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print()
    print(
        "💾 SAVED:",
        OUTPUT_FILE
    )

    print(
        "TOTAL ROWS:",
        len(df)
    )


# ============================================================
# LOAD PREVIOUS RESULTS
# ============================================================

def load_existing_results():

    path = Path(
        OUTPUT_FILE
    )

    if not path.exists():

        return []

    try:

        df = pd.read_excel(
            OUTPUT_FILE
        )

        required = [
            "Page Name",
            "Post URL",
            "Comment",
            "Profile URL",
        ]

        if not all(
            x in df.columns
            for x in required
        ):

            return []

        df = df.fillna("")

        results = df[
            required
        ].to_dict(
            "records"
        )

        print()
        print(
            f"Existing output loaded: "
            f"{len(results)} rows"
        )

        return results

    except Exception as e:

        print(
            "Could not load existing output:",
            e
        )

        return []


# ============================================================
# CHECK IF POST ALREADY PROCESSED
# ============================================================

def post_already_processed(
    results,
    company,
    post_url
):

    company = normalize_text(
        company
    ).lower()

    post_url = normalize_text(
        post_url
    ).lower()

    for item in results:

        existing_company = normalize_text(
            item.get(
                "Page Name",
                ""
            )
        ).lower()

        existing_post = normalize_text(
            item.get(
                "Post URL",
                ""
            )
        ).lower()

        if (
            existing_company == company
            and
            existing_post == post_url
        ):

            return True

    return False


# ============================================================
# MAIN
# ============================================================

async def main():

    # ========================================================
    # LOAD COMPANIES
    # ========================================================

    companies = load_companies()

    # ========================================================
    # LOAD OLD RESULTS
    # ========================================================

    all_results = (
        load_existing_results()
    )

    print()
    print("=" * 70)
    print("FINAL RUN")
    print("=" * 70)

    print(
        f"Companies: {len(companies)}"
    )

    print(
        f"Max posts/page: {MAX_POSTS_PER_PAGE}"
    )

    print(
        f"Theoretical maximum posts: "
        f"{len(companies) * MAX_POSTS_PER_PAGE}"
    )

    print(
        f"Existing result rows: "
        f"{len(all_results)}"
    )

    print(
        "=" * 70
    )

    # ========================================================
    # PLAYWRIGHT
    # ========================================================

    async with async_playwright() as p:

        print()
        print(
            "Starting Chromium..."
        )

        context = (
            await p.chromium.launch_persistent_context(

                user_data_dir=str(
                    PROFILE_DIR
                ),

                headless=HEADLESS,

                viewport={
                    "width": 1366,
                    "height": 900,
                },

                locale="en-US",

                timezone_id="Africa/Cairo",

                args=[
                    "--disable-blink-features=AutomationControlled"
                ],
            )
        )

        # ====================================================
        # LOGIN PAGE
        # ====================================================

        if context.pages:

            login_page = (
                context.pages[0]
            )

        else:

            login_page = (
                await context.new_page()
            )

        print()
        print(
            "Opening Instagram..."
        )

        try:

            await login_page.goto(
                "https://www.instagram.com/",
                wait_until="domcontentloaded",
                timeout=60000
            )

        except Exception as e:

            print(
                "Instagram navigation warning:",
                e
            )

        await safe_wait(
            login_page,
            6000
        )

        # ====================================================
        # LOGIN
        # ====================================================

        print()
        print("=" * 70)
        print("LOGIN CHECK")
        print("=" * 70)

        print()
        print(
            "لو Instagram طالب Login:"
        )

        print(
            "1. سجل الدخول يدويًا."
        )

        print(
            "2. لو طلب Verification Code دخله."
        )

        print(
            "3. خلّص أي Security Check."
        )

        print(
            "4. تأكد إن Instagram فتح طبيعي."
        )

        print()
        print(
            "لو الحساب داخل بالفعل،"
        )

        print(
            "اضغط ENTER مباشرة."
        )

        input(
            "\nاضغط ENTER للبدء..."
        )

        # ====================================================
        # PROCESS COMPANIES
        # ====================================================

        for company_index, row in companies.iterrows():

            company_number = (
                company_index + 1
            )

            company = normalize_text(
                row["Company Name"]
            )

            profile_url = normalize_text(
                row["Instagram"]
            )

            print()
            print()
            print(
                "#" * 75
            )

            print(
                f"PAGE {company_number}/"
                f"{len(companies)}"
            )

            print(
                f"Company: {company}"
            )

            print(
                f"Instagram: {profile_url}"
            )

            print(
                "#" * 75
            )

            profile_page = None

            try:

                # =================================================
                # OPEN PROFILE
                # =================================================

                profile_page = (
                    await open_profile(
                        context,
                        profile_url,
                        company
                    )
                )

                await safe_wait(
                    profile_page,
                    WAIT_AFTER_PROFILE * 1000
                )

                # =================================================
                # GET POSTS
                # =================================================

                post_links = (
                    await get_post_links(
                        profile_page
                    )
                )

                print()
                print(
                    f"PAGE {company_number}: "
                    f"{len(post_links)} posts selected"
                )

                if not post_links:

                    print(
                        "⚠️ No posts found."
                    )

                    continue

                # =================================================
                # PROCESS POSTS
                # =================================================

                for post_number, post_url in enumerate(
                    post_links,
                    start=1
                ):

                    print()
                    print(
                        "-" * 70
                    )

                    print(
                        f"PAGE {company_number}/"
                        f"{len(companies)} | "
                        f"POST {post_number}/"
                        f"{len(post_links)}"
                    )

                    print(
                        f"POST URL: {post_url}"
                    )

                    print(
                        "-" * 70
                    )

                    # =================================================
                    # SKIP ALREADY PROCESSED POST
                    # =================================================

                    if post_already_processed(
                        all_results,
                        company,
                        post_url
                    ):

                        print(
                            "Already exists in output."
                        )

                        print(
                            "Skipping..."
                        )

                        continue

                    post_page = None

                    try:

                        # =================================================
                        # OPEN POST
                        # =================================================

                        post_page = (
                            await open_post(
                                context,
                                post_url
                            )
                        )

                        # =================================================
                        # EXTRACT
                        # =================================================

                        post_results = (
                            await extract_comments_and_profiles(
                                post_page,
                                post_url,
                                company
                            )
                        )

                        # =================================================
                        # ADD RESULTS
                        # =================================================

                        all_results.extend(
                            post_results
                        )

                        print()
                        print(
                            "TOTAL COLLECTED SO FAR:",
                            len(all_results)
                        )

                        # =================================================
                        # SAVE
                        # =================================================

                        save_results(
                            all_results
                        )

                    except Exception as e:

                        print()
                        print(
                            "❌ ERROR PROCESSING POST"
                        )

                        print(
                            type(e).__name__
                        )

                        print(
                            str(e)
                        )

                    finally:

                        if post_page:

                            try:

                                await post_page.close()

                            except Exception:
                                pass

                    await safe_wait(
                        profile_page,
                        1200
                    )

            except Exception as e:

                print()
                print(
                    "❌ ERROR PROCESSING PAGE"
                )

                print(
                    type(e).__name__
                )

                print(
                    str(e)
                )

            finally:

                if profile_page:

                    try:

                        await profile_page.close()

                    except Exception:
                        pass

        # ========================================================
        # FINAL SAVE
        # ========================================================

        save_results(
            all_results
        )

        # ========================================================
        # FINAL REPORT
        # ========================================================

        print()
        print()
        print(
            "=" * 75
        )

        print(
            "🎉 FINISHED"
        )

        print(
            "=" * 75
        )

        print()
        print(
            f"Companies processed: "
            f"{len(companies)}"
        )

        print(
            f"Maximum posts per company: "
            f"{MAX_POSTS_PER_PAGE}"
        )

        print(
            f"Theoretical maximum posts: "
            f"{len(companies) * MAX_POSTS_PER_PAGE}"
        )

        print(
            f"Collected rows: "
            f"{len(all_results)}"
        )

        print()
        print(
            "Output:"
        )

        print(
            OUTPUT_FILE
        )

        print()
        print(
            "Columns:"
        )

        print(
            "Page Name"
        )

        print(
            "Post URL"
        )

        print(
            "Comment"
        )

        print(
            "Profile URL"
        )

        print()
        print(
            "Browser profile:"
        )

        print(
            str(PROFILE_DIR)
        )

        print()
        print(
            "=" * 75
        )

        print()
        print(
            "مهم:"
        )

        print(
            "- الكود لا يخترع Profile URL."
        )

        print(
            "- لو Instagram أظهر رابط صاحب التعليق، سيتم حفظه."
        )

        print(
            "- الكود لا يعتبر 5m أو 2h أو 3w Username."
        )

        print(
            "- الكود لا يعتبر أسماء صفحات Instagram Username للمعلق."
        )

        print(
            "- يتم الحفظ بعد كل Post."
        )

        print(
            "- لو البرنامج وقف، يمكنه استخدام ملف النتائج الموجود."
        )

        print(
            "- الناتج النهائي يحتوي على 4 أعمدة فقط."
        )

        print()

        # إبقاء المتصفح مفتوحًا
        try:

            input(
                "اضغط ENTER لإغلاق المتصفح..."
            )

        except KeyboardInterrupt:

            pass

        try:

            await context.close()

        except Exception:

            pass


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
    
