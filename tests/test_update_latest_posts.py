import unittest
from unittest.mock import patch

import requests

import update_latest_posts


class UpdateLatestPostsTests(unittest.TestCase):
    def test_fetch_user_updates_wraps_requests_http_errors(self):
        with patch.object(
            update_latest_posts,
            "fetch",
            side_effect=requests.exceptions.HTTPError("403 Client Error: Forbidden"),
        ):
            with self.assertRaisesRegex(RuntimeError, "Failed fetching test-user"):
                update_latest_posts.fetch_user_updates(
                    uid="123",
                    username="test-user",
                    existing=[],
                    seen=set(),
                    cookie="cookie",
                    delay=0,
                    timeout=1,
                    overlap=0,
                    page_limit=1,
                )


if __name__ == "__main__":
    unittest.main()
