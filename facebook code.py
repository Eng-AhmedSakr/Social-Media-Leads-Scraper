import asyncio
import re
import pandas as pd
from pathlib import Path
from playwright.async_api import async_playwright


# ============================================================
# SETTINGS
# ============================================================

EXCEL_FILE = r"D:\Data Analysis\Works\New folder\Facebook.xlsx"

OUTPUT_FILE = r"D:\Data Analysis\Works\New folder\Facebook_Comments_Results.xlsx"

# ============================================================
# TEST / RUN SETTINGS
# ============================================================

# أول 5 صفحات للتجربة
MAX_PAGES = 50

# أقصى عدد بوستات يتم فحصها في الصفحة
MAX_POSTS_TO_SCAN = 50

# عدد البوستات التي يجب أن تحتوي على تعليقات عربية
# بمجرد الوصول لهذا الرقم يتوقف البحث في الصفحة
TARGET_ARABIC_POSTS = 10

# عدد مرات Scroll أثناء البحث عن البوستات
MAX_SCROLLS = 100

# عدد مرات محاولة تحميل التعليقات داخل البوست
MAX_COMMENT_ROUNDS = 15

HEADLESS = False

PROFILE_DIR = Path(
    r"D:\Data Analysis\Works\New folder\facebook_profile"
)


# ============================================================
# ARABIC FILTER
# ============================================================

ARABIC_RE = re.compile(
    r"[\u0600-\u06FF]"
)


def normalize_text(text):

    if text is None:
        return ""

    text = str(text)

    text = text.replace(
        "\u200b",
        " "
    )

    text = text.replace(
        "\xa0",
        " "
    )

    return re.sub(
        r"\s+",
        " ",
        text
    ).strip()


def contains_arabic(text):

    return bool(
        ARABIC_RE.search(
            normalize_text(text)
        )
    )


# ============================================================
# FACEBOOK UI TEXT
# ============================================================

IGNORED_TEXT = {

    "Like",
    "Comment",
    "Share",
    "Send",
    "Reply",
    "Replies",

    "See more",
    "See less",

    "Most relevant",
    "All comments",
    "Newest",
    "All Posts",

    "Write a comment...",
    "Write a comment",
    "Comment as",

    "Translate",
    "See translation",

    "Follow",
    "Following",
    "Message",

    "Log in",
    "Log In",
    "Sign Up",

    "Photos",
    "Videos",
    "Posts",
    "Reels",

    "Facebook",
    "Meta",

    "Home",
    "About",
    "Privacy",
    "Terms",

    "Edited",

    "View more comments",
    "View previous comments",
    "View more replies",

    "Most relevant",
    "Newest",
    "All comments",
}


def is_ignored_text(text):

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

    # Pure numbers
    if re.fullmatch(
        r"[\d\s.,:+\-()]+",
        text
    ):
        return True

    # Counters such as:
    # 10 likes
    # 5 comments
    # 3 replies
    if re.fullmatch(
        r"\d[\d,.\s]*\s+"
        r"(likes?|comments?|replies?)",
        text,
        re.I
    ):
        return True

    # Time labels
    if re.fullmatch(
        r"\d+\s*"
        r"(m|min|mins|h|hr|hrs|d|days|w|weeks|mo|months|y|years)"
        r"(?:\s+ago)?",
        text,
        re.I
    ):
        return True

    # Dates
    if re.fullmatch(
        r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",
        text
    ):
        return True

    # URL
    if re.fullmatch(
        r"https?://\S+",
        text,
        re.I
    ):
        return True

    return False


# ============================================================
# VALID ARABIC COMMENT
# ============================================================

def looks_like_arabic_comment(text):

    text = normalize_text(text)

    if not text:
        return False

    if not contains_arabic(text):
        return False

    if is_ignored_text(text):
        return False

    # Too short
    if len(text) < 2:
        return False

    # Too long = probably page text / post text
    if len(text) > 500:
        return False

    # Pure mention
    if re.fullmatch(
        r"@[\w.\-]+",
        text
    ):
        return False

    # Common Facebook UI
    bad_phrases = {

        "عرض المزيد",
        "عرض المزيد من التعليقات",
        "عرض الردود",
        "إظهار المزيد",
        "ترجمة",
        "مشاهدة الترجمة",
        "أعجبني",
        "تعليق",
        "مشاركة",
        "رد",
        "متابعة",
        "رسالة",

    }

    if text.strip() in bad_phrases:
        return False

    return True


# ============================================================
# SAFE WAIT
# ============================================================

async def safe_wait(
    page,
    ms=1500
):

    try:

        await page.wait_for_timeout(
            ms
        )

    except Exception:

        pass


# ============================================================
# LOAD EXCEL
# ============================================================

def load_pages():

    print(
        "\n"
        + "=" * 70
    )

    print(
        "READING FACEBOOK EXCEL"
    )

    print(
        "=" * 70
    )

    df = pd.read_excel(
        EXCEL_FILE
    )

    print(
        "\nDetected columns:"
    )

    for col in df.columns:

        print(
            " -",
            col
        )

    required = [
        "Brand",
        "Facebook URL",
    ]

    for col in required:

        if col not in df.columns:

            raise ValueError(
                f"\n❌ Missing required column: {col}\n"
                f"Detected columns: {list(df.columns)}"
            )

    df = df[
        required
    ].copy()

    df = df.dropna(
        subset=[
            "Facebook URL"
        ]
    )

    df["Brand"] = (
        df["Brand"]
        .astype(str)
        .str.strip()
    )

    df["Facebook URL"] = (
        df["Facebook URL"]
        .astype(str)
        .str.strip()
    )

    df = df[
        df["Facebook URL"]
        .str.startswith(
            "http"
        )
    ]

    df = df.head(
        MAX_PAGES
    )

    print(
        "\nPages loaded:",
        len(df)
    )

    print(
        "Target Arabic-comment posts:",
        TARGET_ARABIC_POSTS
    )

    print(
        "Maximum posts to scan/page:",
        MAX_POSTS_TO_SCAN
    )

    return df


# ============================================================
# CLEAN FACEBOOK URL
# ============================================================

def clean_url(url):

    if not url:
        return ""

    url = url.split("?")[0]

    return url.rstrip("/") + "/"


# ============================================================
# IS POST URL?
# ============================================================

def is_post_url(url):

    if not url:
        return False

    low = url.lower()

    patterns = [

        "/posts/",
        "/reel/",
        "/videos/",
        "story_fbid=",
        "/permalink/",
    ]

    return any(
        x in low
        for x in patterns
    )


# ============================================================
# GET POST LINKS FROM PAGE
# ============================================================

async def collect_post_links(
    page,
    target_count=MAX_POSTS_TO_SCAN
):

    print(
        "\n"
        + "=" * 70
    )

    print(
        "[POST SEARCH]"
    )

    print(
        "Searching deeply for Facebook posts..."
    )

    print(
        "Maximum posts to scan:",
        target_count
    )

    print(
        "=" * 70
    )

    seen = set()

    stable_rounds = 0

    previous_count = 0

    for scroll_number in range(
        1,
        MAX_SCROLLS + 1
    ):

        # ----------------------------------------------------
        # Collect current links
        # ----------------------------------------------------

        try:

            links = await page.locator(
                "a"
            ).evaluate_all(
                """
                elements => {

                    return elements
                        .map(a => a.href)
                        .filter(Boolean);

                }
                """
            )

        except Exception:

            links = []

        for link in links:

            if not is_post_url(
                link
            ):
                continue

            clean = clean_url(
                link
            )

            # Remove weird Facebook tracking
            clean = clean.split(
                "?__cft__"
            )[0]

            clean = clean.split(
                "?__tn__"
            )[0]

            if clean not in seen:

                seen.add(
                    clean
                )

        current_count = len(
            seen
        )

        print(
            f"Scroll {scroll_number}/{MAX_SCROLLS} | "
            f"Discovered: {current_count}"
        )

        # ----------------------------------------------------
        # Enough posts discovered
        # ----------------------------------------------------

        if current_count >= target_count:

            print(
                "\n[OK] Maximum post scan reached."
            )

            break

        # ----------------------------------------------------
        # Detect no new posts
        # ----------------------------------------------------

        if current_count == previous_count:

            stable_rounds += 1

        else:

            stable_rounds = 0

        previous_count = current_count

        # ----------------------------------------------------
        # Scroll
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # If Facebook stopped loading
        # ----------------------------------------------------

        if stable_rounds >= 12:

            print(
                "\n[INFO] No new posts detected."
            )

            print(
                "Trying additional scrolling..."
            )

            # Small extra scroll
            try:

                await page.mouse.wheel(
                    0,
                    6000
                )

            except Exception:

                pass

            await safe_wait(
                page,
                2500
            )

            # Reset
            stable_rounds = 0

    result = list(
        seen
    )

    result = result[
        :target_count
    ]

    print(
        "\n[POST SEARCH COMPLETE]"
    )

    print(
        "Posts discovered:",
        len(result)
    )

    return result


# ============================================================
# OPEN POST
# ============================================================

async def open_post(
    context,
    url
):

    post_page = await context.new_page()

    try:

        await post_page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=60000
        )

    except Exception as e:

        print(
            "Post navigation warning:",
            type(e).__name__,
            str(e)
        )

    await safe_wait(
        post_page,
        4000
    )

    return post_page


# ============================================================
# LOAD COMMENTS
# ============================================================

async def load_comments(
    page
):

    print(
        "\n[COMMENTS] Loading comments..."
    )

    buttons = [

        "View more comments",
        "View previous comments",
        "View more replies",
        "View previous replies",

        "Load more comments",
        "Load more replies",

        "See more comments",
        "See more replies",

    ]

    total_clicks = 0

    for round_number in range(
        1,
        MAX_COMMENT_ROUNDS + 1
    ):

        clicked_this_round = 0

        for text in buttons:

            try:

                locator = page.get_by_text(
                    text,
                    exact=False
                )

                count = await locator.count()

                for i in range(
                    min(
                        count,
                        20
                    )
                ):

                    try:

                        element = locator.nth(
                            i
                        )

                        if await element.is_visible():

                            await element.click(
                                timeout=2500
                            )

                            clicked_this_round += 1
                            total_clicks += 1

                            await safe_wait(
                                page,
                                700
                            )

                    except Exception:

                        pass

            except Exception:

                pass

        print(
            f"  Comment round "
            f"{round_number}/{MAX_COMMENT_ROUNDS} "
            f"| Clicked: "
            f"{clicked_this_round}"
        )

        # Scroll
        try:

            await page.mouse.wheel(
                0,
                1400
            )

        except Exception:

            pass

        await safe_wait(
            page,
            900
        )

        # If nothing clicked several times,
        # still continue scrolling because Facebook
        # sometimes loads comments automatically.
        if (
            clicked_this_round == 0
            and round_number >= 8
        ):

            try:

                await page.mouse.wheel(
                    0,
                    2500
                )

            except Exception:

                pass

            await safe_wait(
                page,
                1200
            )

    print(
        "[COMMENTS] Total clicks:",
        total_clicks
    )


# ============================================================
# EXTRACT PROFILE URL
# ============================================================

def extract_facebook_profile(
    html
):

    if not html:
        return None

    urls = re.findall(
        r'href=["\']([^"\']+)["\']',
        html
    )

    bad_words = {

        "facebook",
        "privacy",
        "terms",
        "help",
        "settings",
        "login",
        "watch",
        "reels",
        "videos",
        "photos",
        "groups",
        "marketplace",
        "events",
        "pages",
        "gaming",
        "friends",
        "notifications",

    }

    for url in urls:

        url = url.replace(
            "&amp;",
            "&"
        )

        if url.startswith("/"):
            full = (
                "https://www.facebook.com"
                + url
            )
        else:
            full = url

        low = full.lower()

        # Ignore post/reel/video links
        if (
            "/posts/" in low
            or "/reel/" in low
            or "/videos/" in low
            or "story_fbid=" in low
        ):
            continue

        # Profile URLs generally look like:
        # /username/
        # /profile.php?id=...
        if "/profile.php" in low:

            return full.split("?")[0]

        match = re.search(
            r"facebook\.com/([^/?#]+)",
            full,
            re.I
        )

        if not match:
            continue

        username = match.group(
            1
        )

        if username.lower() in bad_words:
            continue

        if username.startswith(
            "photo"
        ):
            continue

        if username.startswith(
            "story"
        ):
            continue

        if username.startswith(
            "posts"
        ):
            continue

        return (
            "https://www.facebook.com/"
            + username
            + "/"
        )

    return None


# ============================================================
# EXTRACT NAME
# ============================================================

def extract_name_from_block(
    text,
    profile_url
):

    lines = [

        normalize_text(x)

        for x in text.splitlines()

        if normalize_text(x)

    ]

    if not lines:
        return ""

    username = ""

    if profile_url:

        match = re.search(
            r"facebook\.com/([^/?#]+)",
            profile_url,
            re.I
        )

        if match:

            username = (
                match.group(1)
                .lower()
            )

    for line in lines:

        if is_ignored_text(
            line
        ):
            continue

        if contains_arabic(
            line
        ):

            # This could be comment,
            # not necessarily name.
            continue

        if len(line) > 100:
            continue

        # Skip obvious UI
        low = line.lower()

        if low in {
            "like",
            "comment",
            "share",
            "reply",
            "follow",
            "following",
            "message",
            "translate",
            "edited",
        }:
            continue

        if username and (
            line.lower()
            == username
        ):
            continue

        # Names often contain letters/spaces
        if re.fullmatch(
            r"[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ .'\-]{1,80}",
            line
        ):
            return line

    return ""


# ============================================================
# EXTRACT COMMENTS
# ============================================================

async def extract_arabic_comments(
    page,
    post_url,
    brand,
    page_url
):

    await load_comments(
        page
    )

    print(
        "\n[EXTRACTING]"
    )

    results = []

    seen = set()

    # --------------------------------------------------------
    # First method:
    # Facebook comment blocks
    # --------------------------------------------------------

    selectors = [

        'div[role="article"]',

        'div[data-pagelet*="Comment"]',

        'div[aria-label*="Comment"]',

    ]

    blocks = []

    for selector in selectors:

        try:

            data = await page.locator(
                selector
            ).evaluate_all(
                """
                elements => elements.map(
                    el => ({
                        text: el.innerText || "",
                        html: el.outerHTML || ""
                    })
                )
                """
            )

            if data:

                blocks.extend(
                    data
                )

        except Exception:

            pass

    print(
        "[INFO] Candidate comment blocks:",
        len(blocks)
    )

    # --------------------------------------------------------
    # Process blocks
    # --------------------------------------------------------

    for block in blocks:

        text = normalize_text(
            block.get(
                "text",
                ""
            )
        )

        html = block.get(
            "html",
            ""
        )

        if not text:
            continue

        profile_url = (
            extract_facebook_profile(
                html
            )
        )

        name = (
            extract_name_from_block(
                text,
                profile_url
            )
        )

        lines = [

            normalize_text(
                x
            )

            for x in text.splitlines()

            if normalize_text(x)

        ]

        comment = ""

        for line in lines:

            # Must be Arabic
            if not contains_arabic(
                line
            ):
                continue

            if not looks_like_arabic_comment(
                line
            ):
                continue

            # Avoid using the entire post text
            if len(line) > 500:
                continue

            comment = line

            break

        if not comment:

            continue

        key = (
            page_url,
            post_url,
            profile_url or "",
            comment,
        )

        if key in seen:

            continue

        seen.add(
            key
        )

        results.append(
            {
                "Source Page Name": brand,
                "Source Page URL": page_url,
                "Source Post URL": post_url,
                "Comment": comment,
                "Full Name": name,
                "Facebook URL": profile_url or "",
                "Phone Number": "",
            }
        )

    # --------------------------------------------------------
    # FALLBACK:
    # Search visible page text for Arabic
    # --------------------------------------------------------

    if not results:

        print(
            "[INFO] No Arabic comments found "
            "from comment blocks."
        )

        print(
            "[INFO] Trying page-text fallback..."
        )

        try:

            body = await page.locator(
                "body"
            ).inner_text(
                timeout=15000
            )

            lines = [

                normalize_text(
                    x
                )

                for x in body.splitlines()

                if normalize_text(x)

            ]

            for line in lines:

                if not contains_arabic(
                    line
                ):
                    continue

                if not looks_like_arabic_comment(
                    line
                ):
                    continue

                # Ignore huge text
                if len(line) > 500:
                    continue

                key = (
                    page_url,
                    post_url,
                    "",
                    line,
                )

                if key in seen:
                    continue

                seen.add(
                    key
                )

                results.append(
                    {
                        "Source Page Name": brand,
                        "Source Page URL": page_url,
                        "Source Post URL": post_url,
                        "Comment": line,
                        "Full Name": "",
                        "Facebook URL": "",
                        "Phone Number": "",
                    }
                )

                # Don't flood from fallback
                if len(results) >= 20:
                    break

        except Exception as e:

            print(
                "Fallback error:",
                type(e).__name__,
                str(e)
            )

    # --------------------------------------------------------
    # Final validation
    # --------------------------------------------------------

    final = []

    for item in results:

        comment = normalize_text(
            item["Comment"]
        )

        if not looks_like_arabic_comment(
            comment
        ):
            continue

        item["Comment"] = comment

        final.append(
            item
        )

    print(
        "\n[ARABIC COMMENTS FOUND]:",
        len(final)
    )

    for item in final[:10]:

        print(
            " ",
            item["Full Name"],
            "=>",
            item["Comment"]
        )

    return final


# ============================================================
# EXTRACT PUBLIC PHONE NUMBER
# ============================================================

def extract_phone_numbers(
    text
):

    if not text:
        return []

    # Egyptian phone formats
    patterns = [

        r"(?:\+20|0020)\s*1[0125]\s*\d{8}",

        r"01[0125]\s*\d{8}",

        r"\+20\s*1[0125]\s*\d{8}",

    ]

    found = []

    for pattern in patterns:

        matches = re.findall(
            pattern,
            text
        )

        for number in matches:

            clean = re.sub(
                r"[^\d+]",
                "",
                number
            )

            if clean not in found:

                found.append(
                    clean
                )

    return found


async def extract_public_phone(
    context,
    profile_url
):

    if not profile_url:

        return ""

    profile_page = None

    try:

        profile_page = (
            await context.new_page()
        )

        await profile_page.goto(
            profile_url,
            wait_until="domcontentloaded",
            timeout=30000
        )

        await safe_wait(
            profile_page,
            2500
        )

        # ----------------------------------------------------
        # First: tel links
        # ----------------------------------------------------

        try:

            tel_links = await profile_page.locator(
                'a[href^="tel:"]'
            ).evaluate_all(
                """
                elements =>
                    elements.map(
                        e => e.href
                    )
                """
            )

        except Exception:

            tel_links = []

        for link in tel_links:

            number = link.replace(
                "tel:",
                ""
            )

            numbers = extract_phone_numbers(
                number
            )

            if numbers:

                return numbers[0]

        # ----------------------------------------------------
        # Second: visible page text
        # ----------------------------------------------------

        try:

            body = await profile_page.locator(
                "body"
            ).inner_text(
                timeout=10000
            )

        except Exception:

            body = ""

        numbers = extract_phone_numbers(
            body
        )

        if numbers:

            return numbers[0]

    except Exception as e:

        print(
            "Phone extraction warning:",
            type(e).__name__
        )

    finally:

        if profile_page:

            try:

                await profile_page.close()

            except Exception:

                pass

    return ""


# ============================================================
# ADD PUBLIC PHONE NUMBERS
# ============================================================

async def enrich_phone_numbers(
    context,
    results
):

    if not results:

        return results

    print(
        "\n[PHONE CHECK]"
    )

    checked_profiles = {}

    for item in results:

        profile_url = (
            item.get(
                "Facebook URL",
                ""
            )
        )

        if not profile_url:

            continue

        if profile_url in checked_profiles:

            item["Phone Number"] = (
                checked_profiles[
                    profile_url
                ]
            )

            continue

        print(
            "Checking public phone:",
            profile_url
        )

        phone = await extract_public_phone(
            context,
            profile_url
        )

        checked_profiles[
            profile_url
        ] = phone

        item["Phone Number"] = phone

        if phone:

            print(
                "  [PHONE FOUND]",
                phone
            )

        else:

            print(
                "  [NO PUBLIC PHONE]"
            )

    return results


# ============================================================
# SAVE EXCEL
# ============================================================

def save_results(
    results
):

    columns = [

        "Source Page Name",
        "Source Page URL",
        "Source Post URL",
        "Comment",
        "Full Name",
        "Facebook URL",
        "Phone Number",

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

        df = df.drop_duplicates(
            subset=[
                "Source Page URL",
                "Source Post URL",
                "Comment",
                "Facebook URL",
            ]
        )

        # Final Arabic-only filter
        df = df[
            df["Comment"].apply(
                contains_arabic
            )
        ]

    else:

        df = pd.DataFrame(
            columns=columns
        )

    df.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print(
        "\n"
        + "=" * 70
    )

    print(
        "EXCEL SAVED"
    )

    print(
        "=" * 70
    )

    print(
        "File:",
        OUTPUT_FILE
    )

    print(
        "Rows:",
        len(df)
    )

    print(
        "=" * 70
    )


# ============================================================
# CHECK FACEBOOK LOGIN
# ============================================================

async def check_login(
    context
):

    if context.pages:

        page = context.pages[0]

    else:

        page = await context.new_page()

    try:

        await page.goto(
            "https://www.facebook.com/",
            wait_until="domcontentloaded",
            timeout=60000
        )

    except Exception as e:

        print(
            "Facebook opening warning:",
            e
        )

    await safe_wait(
        page,
        5000
    )

    print(
        "\n"
        + "=" * 70
    )

    print(
        "CHECKING FACEBOOK LOGIN"
    )

    print(
        "=" * 70
    )

    print(
        "\nCurrent URL:",
        page.url
    )

    try:

        print(
            "Title:",
            await page.title()
        )

    except Exception:

        pass

    print(
        """
لو Facebook طالب Login:

1. سجل الدخول يدويًا.
2. خلص أي Security Check.
3. تأكد إن الحساب دخل Facebook.
4. اضغط Enter.

لو الحساب داخل بالفعل:
اضغط Enter مباشرة.
"""
    )

    input(
        "\nاضغط Enter للبدء..."
    )

    return page


# ============================================================
# OPEN PAGE
# ============================================================

async def open_brand_page(
    context,
    url,
    brand
):

    print(
        "\n"
        + "#" * 70
    )

    print(
        "BRAND:",
        brand
    )

    print(
        "URL:",
        url
    )

    print(
        "#" * 70
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
            "Page navigation warning:",
            type(e).__name__,
            str(e)
        )

    await safe_wait(
        page,
        4000
    )

    return page


# ============================================================
# PROCESS ONE BRAND
# ============================================================

async def process_brand(
    context,
    brand,
    page_url,
    page_number,
    total_pages,
    all_results
):

    print(
        "\n"
        + "#" * 70
    )

    print(
        f"PAGE {page_number}/{total_pages}"
    )

    print(
        "Brand:",
        brand
    )

    print(
        "URL:",
        page_url
    )

    print(
        "#" * 70
    )

    profile_page = None

    try:

        # ----------------------------------------------------
        # Open page
        # ----------------------------------------------------

        profile_page = await open_brand_page(
            context,
            page_url,
            brand
        )

        # ----------------------------------------------------
        # Collect lots of posts
        # ----------------------------------------------------

        post_links = await collect_post_links(
            profile_page,
            MAX_POSTS_TO_SCAN
        )

        if not post_links:

            print(
                "\n❌ No posts found."
            )

            return all_results

        print(
            "\nTotal candidate posts:",
            len(post_links)
        )

        # ----------------------------------------------------
        # SMART POST LOOP
        # ----------------------------------------------------

        checked_posts = 0

        arabic_posts = 0

        for post_url in post_links:

            # IMPORTANT:
            # This is the smart stopping mechanism.
            # We do NOT care how many posts were checked.
            # We care how many posts actually contain
            # Arabic comments.

            if arabic_posts >= TARGET_ARABIC_POSTS:

                print(
                    "\n"
                    + "=" * 70
                )

                print(
                    "[TARGET REACHED]"
                )

                print(
                    f"Found {arabic_posts} posts "
                    f"with Arabic comments."
                )

                print(
                    "Stopping this page immediately."
                )

                print(
                    "=" * 70
                )

                break

            if checked_posts >= MAX_POSTS_TO_SCAN:

                break

            checked_posts += 1

            print(
                "\n"
                + "-" * 70
            )

            print(
                "[CHECKING POST]"
            )

            print(
                f"Checked post: "
                f"{checked_posts}/"
                f"{MAX_POSTS_TO_SCAN}"
            )

            print(
                f"Arabic posts: "
                f"{arabic_posts}/"
                f"{TARGET_ARABIC_POSTS}"
            )

            print(
                post_url
            )

            print(
                "-" * 70
            )

            post_page = None

            try:

                # ------------------------------------------------
                # Open post
                # ------------------------------------------------

                post_page = await open_post(
                    context,
                    post_url
                )

                # ------------------------------------------------
                # Extract Arabic comments
                # ------------------------------------------------

                post_results = await extract_arabic_comments(
                    post_page,
                    post_url,
                    brand,
                    page_url
                )

                # ------------------------------------------------
                # THE IMPORTANT PART
                #
                # Post counts ONLY if Arabic comments exist.
                # ------------------------------------------------

                if post_results:

                    arabic_posts += 1

                    print(
                        "\n"
                        + "=" * 60
                    )

                    print(
                        "✅ ARABIC COMMENTS FOUND"
                    )

                    print(
                        f"Arabic posts: "
                        f"{arabic_posts}/"
                        f"{TARGET_ARABIC_POSTS}"
                    )

                    print(
                        "➡️ POST COUNTED"
                    )

                    print(
                        "=" * 60
                    )

                    # ------------------------------------------------
                    # Check public phone numbers
                    # ------------------------------------------------

                    post_results = await enrich_phone_numbers(
                        context,
                        post_results
                    )

                    # Add results
                    all_results.extend(
                        post_results
                    )

                    # Save immediately
                    save_results(
                        all_results
                    )

                else:

                    print(
                        "\n"
                        + "-" * 60
                    )

                    print(
                        "❌ NO ARABIC COMMENTS"
                    )

                    print(
                        "➡️ POST NOT COUNTED"
                    )

                    print(
                        "-" * 60
                    )

            except Exception as e:

                print(
                    "\n❌ ERROR PROCESSING POST"
                )

                print(
                    type(e).__name__,
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
                800
            )

        # ----------------------------------------------------
        # PAGE SUMMARY
        # ----------------------------------------------------

        print(
            "\n"
            + "=" * 70
        )

        print(
            f"PAGE FINISHED: {brand}"
        )

        print(
            "=" * 70
        )

        print(
            "Posts checked:",
            checked_posts
        )

        print(
            "Posts with Arabic comments:",
            arabic_posts
        )

        print(
            "Target:",
            TARGET_ARABIC_POSTS
        )

        print(
            "Total collected rows:",
            len(all_results)
        )

        print(
            "=" * 70
        )

    except Exception as e:

        print(
            "\n❌ ERROR PROCESSING PAGE"
        )

        print(
            type(e).__name__,
            str(e)
        )

    finally:

        if profile_page:

            try:

                await profile_page.close()

            except Exception:

                pass

    return all_results


# ============================================================
# MAIN
# ============================================================

async def main():

    # --------------------------------------------------------
    # Load Excel
    # --------------------------------------------------------

    pages = load_pages()

    print(
        "\n"
        + "=" * 70
    )

    print(
        "TEST MODE"
    )

    print(
        "=" * 70
    )

    print(
        "Pages to process:",
        len(pages)
    )

    print(
        "Maximum posts to scan/page:",
        MAX_POSTS_TO_SCAN
    )

    print(
        "Target posts WITH Arabic comments:",
        TARGET_ARABIC_POSTS
    )

    print(
        "Language filter: Arabic only"
    )

    print(
        "=" * 70
    )

    all_results = []

    # --------------------------------------------------------
    # Playwright
    # --------------------------------------------------------

    async with async_playwright() as p:

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

                args=[
                    "--disable-blink-features=AutomationControlled"
                ],
            )
        )

        # ----------------------------------------------------
        # Login
        # ----------------------------------------------------

        await check_login(
            context
        )

        # ----------------------------------------------------
        # Process pages
        # ----------------------------------------------------

        total_pages = len(
            pages
        )

        for page_number, (
            index,
            row
        ) in enumerate(
            pages.iterrows(),
            start=1
        ):

            brand = normalize_text(
                row["Brand"]
            )

            page_url = normalize_text(
                row["Facebook URL"]
            )

            if not brand:
                brand = "Unknown"

            if not page_url:
                continue

            all_results = await process_brand(
                context,
                brand,
                page_url,
                page_number,
                total_pages,
                all_results
            )

            # Save after every page
            save_results(
                all_results
            )

        # ----------------------------------------------------
        # Final save
        # ----------------------------------------------------

        save_results(
            all_results
        )

        print(
            "\n"
            + "=" * 70
        )

        print(
            "FINAL RESULTS"
        )

        print(
            "=" * 70
        )

        print(
            "Pages processed:",
            total_pages
        )

        print(
            "Target Arabic-comment posts/page:",
            TARGET_ARABIC_POSTS
        )

        print(
            "Maximum posts scanned/page:",
            MAX_POSTS_TO_SCAN
        )

        print(
            "Total collected rows:",
            len(all_results)
        )

        print(
            "Output:",
            OUTPUT_FILE
        )

        print(
            "=" * 70
        )

        print(
            """
ملاحظات:

1. الكود يعتمد على ملف Facebook.xlsx.

2. الأعمدة المطلوبة في ملف الإدخال:
   Brand
   Facebook URL

3. البرنامج يبحث في البوستات بشكل تدريجي.

4. البوست لا يتحسب ضمن الـ10 إلا إذا تم العثور
   فعليًا على تعليق عربي.

5. لو البوست لا يحتوي على تعليق عربي:
   يتم تجاهله والانتقال للبوست التالي.

6. البرنامج يمكنه فحص حتى 200 بوست لكل صفحة.

7. بمجرد العثور على 10 بوستات بها تعليقات عربية:
   يتوقف البحث في الصفحة فورًا.

8. Comment = التعليق العربي.

9. Source Post URL = رابط البوست.

10. Source Page Name = اسم الصفحة.

11. Source Page URL = رابط صفحة الشركة.

12. Facebook URL = رابط بروفايل المعلق إذا أمكن
    استخراجه فعليًا من عنصر التعليق.

13. Full Name = الاسم الظاهر فعليًا إذا أمكن
    استخراجه من عنصر التعليق.

14. Phone Number = رقم هاتف عام ظاهر فعليًا
    في بروفايل المعلق إن أمكن الوصول إليه.

15. البرنامج لا يخترع اسمًا أو رابط بروفايل
    أو رقم هاتف.

16. يتم حفظ Excel بعد كل بوست مؤهل،
    لذلك النتائج السابقة لا تضيع لو حصل توقف.

17. الكود يعمل على أول 5 صفحات حاليًا.
    بعد نجاح التجربة غيّر:

        MAX_PAGES = 5

    إلى:

        MAX_PAGES = 100
"""
        )

        # ----------------------------------------------------
        # Keep browser open
        # ----------------------------------------------------

        try:

            while True:

                await asyncio.sleep(
                    3600
                )

        except KeyboardInterrupt:

            print(
                "\nStopping program..."
            )

        finally:

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