import unittest

from vn_address_cleaner.cleaner import AddressCleaner
from vn_address_cleaner.llm import validate_llm_result


class StreetRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cleaner = AddressCleaner()

    def test_reported_bad_street_outputs(self):
        cases = [
            ("193 Đê Long Biên, Phường Ngọc Lâm, Quận Long Biên, Thành Phố Hà Nội", "Đê Long Biên"),
            ("dia chi nhà sõ 3 ngõ 48 Ninh Hiêp Phù dông Hà Nôi 0987202786", ""),
            ("64/54 hòa bih p5 q11 sdt 0772710758", "Hòa Bình"),
            ("218 Âu Thôn Naosari Lộc Nga Bảo Lộc Lâm Đồng", ""),
            ("Số48 đông thôn tuân lề_ tiên dương *đông anh* hà Nội", ""),
            ("50 c.thái.đc.xã.hoà.tú.2.huyện.mỹ. xuyên.tỉnh sóc trăng", ""),
            ("20 cây nguyễn thị tuyết quảng xá vĩnh lâm vĩnh linh quảng trị.đt0914134991", ""),
            ("Thông tin 0966972659 B04-L35 ( Nội thất 5C), An Phú shop villa, Dương Nội, Hà Nội", ""),
            ("19 Ngõ 34 Trần Phú Nguyễn Trãi 2 Sao Đỏ", "Trần Phú"),
            ("Sn 19 Đường Đình 1 Xuân Bách Quang Tiến Sóc Sơn Hà Nội", "Đường Đình 1"),
            ("Ship Đường D8-5A Đồng Sổ Lai Uyên Bàu Bàng Bình Dương Cũ", "Đường D8-5A"),
            ("77 Chi Quan Liên Quan Thạch Thất Hà Nội Ạ", ""),
            ("026 lo s chung cu ngo gia tu p2q10 hcm", ""),
            ("số 5 tú xương mạc đinh chi vũng tàu", "Tú Xương"),
            ("122/155 bến Phú Định p16 q8 hcm", "Bến Phú Định"),
            ("k9 tiệm sửa xe Soz ấp Tây An, Xã Thạnh Mỹ Tây, Huyện Châu Phú, An Giang", ""),
            ("46a đường 30/4 hội xuân , tt tầm vu , châu thành , long an", "Đường 30/4"),
            ("Xóm ba đào xá đào duong an thì Hưng yên", ""),
            ("k5 toa đen xã việt Xuân h Vĩnh tường tỉnh Vĩnh phúc đt", ""),
            ("32 Hung Vuong Thôn Thương Châu Xã Ngũ Phung Huyện Phú Quý Tinh Bình Thuận", "Hùng Vương"),
        ]
        for raw, expected in cases:
            with self.subTest(raw=raw):
                self.assertEqual(self.cleaner.clean(raw).street, expected)

    def test_llm_cannot_return_truncated_single_word_street(self):
        raw = "122/155 bến Phú Định p16 q8 hcm"
        llm_result = {
            "poi": None,
            "house_number": "122/155",
            "street": "Bến",
            "level4": None,
            "ward": "Phường 16",
            "district": "Quận 8",
            "province": "Thành phố Hồ Chí Minh",
            "confidence": 0.95,
            "flags": [],
            "evidence": {
                "poi_span": None,
                "house_number_span": "122/155",
                "street_span": "bến",
                "level4_span": None,
            },
        }
        original = {
            "raw_address": raw,
            "ward": "Phường 16",
            "district": "Quận 8",
            "province": "Thành phố Hồ Chí Minh",
            "flags": ["UNPREFIXED_STREET_FROM_HOUSE_NUMBER"],
            "rule_candidates": {"street": "Bến Phú Định"},
        }
        validated = validate_llm_result(llm_result, original)
        self.assertIsNone(validated["street"])
        self.assertIn("LLM_STREET_FORMAT_REJECTED", validated["flags"])

    def test_internal_row_code_and_residential_project(self):
        raw = "20 16B1 Làng Việt Kiều Châu Âu - Mỗ Lao - Hà Đông - HN"
        result = self.cleaner.clean(raw)
        self.assertEqual(result.street, "16B1")
        self.assertEqual(result.poi, "Khu đô thị Làng Việt Kiều Châu Âu")
        self.assertEqual(result.level4, "")
        self.assertEqual(result.ward, "Phường Mộ Lao")
        self.assertEqual(result.district, "Quận Hà Đông")
        self.assertEqual(result.province, "Thành phố Hà Nội")

        llm_result = {
            "poi": "Khu đô thị Làng Việt Kiều Châu Âu",
            "house_number": "20",
            "street": "16B1",
            "level4": None,
            "ward": result.ward,
            "district": result.district,
            "province": result.province,
            "confidence": 0.95,
            "flags": [],
            "evidence": {
                "poi_span": "Làng Việt Kiều Châu Âu",
                "house_number_span": "20",
                "street_span": "16B1",
                "level4_span": None,
            },
        }
        original = {
            "raw_address": raw,
            "ward": result.ward,
            "district": result.district,
            "province": result.province,
            "flags": list(result.flags),
            "rule_candidates": {
                "poi": result.poi,
                "street": result.street,
                "level4": result.level4,
            },
        }
        validated = validate_llm_result(llm_result, original)
        self.assertEqual(validated["street"], "16B1")
        self.assertEqual(validated["poi"], "Khu Đô Thị Làng Việt Kiều Châu Âu")

    def test_additional_reported_boundaries_and_typos(self):
        cases = [
            (
                "Công Ty Tnhh South & North Vina - Lô G14, Đường G1C, Kcn Quế Võ, Quế Võ, Bắc Ninh",
                "Đường G1C",
            ),
            ("Sn 342 dinh tiên  hoàng", "Đinh Tiên Hoàng"),
            ("Ô4,dãy D,đường võ nguyên giáp", "Đường Võ Nguyên Giáp"),
            ("3/2đ đính tiên hoàng", "Đinh Tiên Hoàng"),
            ("Xóm Đê Hến, Thị Trấn Bá Hiến, Huyện Bình Xuyên, Tỉnh Vĩnh Phúc", ""),
            (
                "Đường Hiếu Lâm 2 Nhánh 1 . Ấp Cây Dừng Xã Hiếu Liêm Bắc Tân Uyên Bình Dương",
                "Đường Hiếu Lâm 2",
            ),
            ("63 Lê thị hoa bình chiểu thủ Đức", "Lê Thị Hoa"),
            ("B7-02 trần anh ashita, ấp 4 trừ văn thố bàu bàng bình dương", ""),
            ("Uh 134 đê la thành đômgs đa hn. 0989842866", "Đê La Thành"),
            ("Số 06D Ql1A Gián Khẩu Gia Trấn Gia Viễn Ninh Bình", "Quốc Lộ 1A"),
            ("02 Lê Lai . Hợp Thành", "Lê Lai"),
            ("Làng Đê Gôh", ""),
        ]
        for raw, expected in cases:
            with self.subTest(raw=raw):
                self.assertEqual(self.cleaner.clean(raw).street, expected)

        project = self.cleaner.clean(
            "B7-02 trần anh ashita, ấp 4 trừ văn thố bàu bàng bình dương"
        )
        self.assertEqual(project.poi, "Trần Anh Ashita")
        self.assertEqual(project.level4, "Ấp 4")

    def test_llm_street_boundary_guards(self):
        def validate(raw, street, evidence, rule, ward="", district="", province=""):
            first_token = raw.split()[0] if raw.split() else ""
            house_number = first_token if any(ch.isdigit() for ch in first_token) else None
            return validate_llm_result(
                {
                    "poi": None,
                    "house_number": house_number,
                    "street": street,
                    "level4": None,
                    "ward": ward,
                    "district": district,
                    "province": province,
                    "confidence": 0.9,
                    "flags": [],
                    "evidence": {
                        "poi_span": None,
                        "house_number_span": house_number,
                        "street_span": evidence,
                        "level4_span": None,
                    },
                },
                {
                    "raw_address": raw,
                    "ward": ward,
                    "district": district,
                    "province": province,
                    "flags": ["UNPREFIXED_STREET_FROM_HOUSE_NUMBER"],
                    "rule_candidates": {"street": rule},
                },
            )

        admin_tail = validate(
            "63 Lê thị hoa bình chiểu thủ Đức",
            "Lê Thị Hoa Bình Chiểu",
            "Lê thị hoa bình chiểu",
            "Lê Thị Hoa",
            "Phường Bình Chiểu",
            "Thành phố Thủ Đức",
            "Thành phố Hồ Chí Minh",
        )
        self.assertEqual(admin_tail["street"], "Lê Thị Hoa")

        dotted = validate(
            "02 Lê Lai . Hợp Thành",
            "Lê Lai Hợp Thành",
            "Lê Lai . Hợp Thành",
            "Lê Lai",
        )
        self.assertEqual(dotted["street"], "Lê Lai")

        project = validate(
            "B7-02 trần anh ashita, ấp 4",
            "Trần Anh Ashita",
            "trần anh ashita",
            "",
        )
        self.assertIsNone(project["street"])

    def test_reported_poi_cleanup_cases(self):
        cases = [
            (
                "Chị Hương 0902102466, Toà Bidv, Số 46 Đường Cao Lỗ, Tổ 2, Thị Trấn Đông Anh, Đông Anh, Hà Nội",
                "Tòa BIDV", "Đường Cao Lỗ", "Tổ 2",
            ),
            (
                "Bidv Cn Mỹ Đình · Địa Chỉ: Tầng 1,2,3 Khu Tổ Hợp Văn Phòng, Trung Tâm Thương Mại Và Chung Cư Cao Cấp Golden Palace, Mễ Trì, Huyện Nam Từ Liêm, Hà Nội",
                "BIDV Chi nhánh Mỹ Đình", "", "",
            ),
            ("Chùa Đội 9. Đồng Bào Gia Xuyên Tp Hải Dương, Tỉnh Hải Dương", "Chùa Đội 9", "", "Đội 9"),
            (
                "Chị Ngọc.650K 6S. Công Ty Bao Bì Hương Sen Sâu Bến Xe Buýt Hoàng Hà Phường Tiền Phong Thành Phố Thái Bình Tỉnh Thái Bình.0865010326",
                "Công ty Bao bì Hương Sen", "", "",
            ),
            ("cổng CTY aac chu mẫu vân dương TP bắc Ninh tỉnh bắc ninh", "Công ty AAC", "", ""),
            (
                "NHÀ 417 - LÔ C CẦU THANG 2 - CHUNG CƯ 43D HỒ VĂN HUÊ P9 PHÚ NHUẬN hcm",
                "Chung cư 43D", "Hồ Văn Huê", "",
            ),
            (
                "cổng chợ Kim SN 156 tdp Kim 1 phường phượng sơn lục ngạn Bắc Giang",
                "Chợ Kim", "", "Tổ dân phố Kim 1",
            ),
            ("cổng chùa thôn chùa vàng minh quang", "Chùa Vàng", "", "Thôn Chùa Vàng"),
            (
                "Bình Minh Binh Mỹ Châu Phu Ăn Giang Chợ Vòng Sáng Cây Dương",
                "Chợ Vòng Sáng Cây Dương", "", "",
            ),
            ("Cty Mỹ Hưng 35B ấp Bắc p5 Mỹ Tho", "Công ty Mỹ Hưng", "", "Ấp Bắc"),
            (
                "Cc Park Kiara Lê Trọng Tấn Dương Nội",
                "Chung cư Park Kiara", "Lê Trọng Tấn", "",
            ),
        ]
        for raw, poi, street, level4 in cases:
            with self.subTest(raw=raw):
                result = self.cleaner.clean(raw)
                self.assertEqual(result.poi, poi)
                self.assertEqual(result.street, street)
                self.assertEqual(result.level4, level4)


if __name__ == "__main__":
    unittest.main()
