

import { PPTXXmlUtils } from './xml.js';
import { PPTXStyleUtils } from './style.js';

export const PPTXTextUtils = (function() {
    var slideFactor = 96 / 914400;
    var fontSizeFactor = 4 / 3.2;
    
    var rtl_langs_array = ["he-IL", "ar-AE", "ar-SA", "dv-MV", "fa-IR","ur-PK"]
    
    var is_first_br = false;

    var dingbat_unicode = [
        {"f": "Webdings", "code": "33",  "unicode": "128375"},
        {"f": "Webdings", "code": "34",  "unicode": "128376"},
        {"f": "Webdings", "code": "35",  "unicode": "128370"},
        {"f": "Webdings", "code": "36",  "unicode": "128374"},
        {"f": "Webdings", "code": "37",  "unicode": "127942"},
        {"f": "Webdings", "code": "38",  "unicode": "127894"},
        {"f": "Webdings", "code": "39",  "unicode": "128391"},
        {"f": "Webdings", "code": "40",  "unicode": "128488"},
        {"f": "Webdings", "code": "41",  "unicode": "128489"},
        {"f": "Webdings", "code": "42",  "unicode": "128496"},
        {"f": "Webdings", "code": "43",  "unicode": "128497"},
        {"f": "Webdings", "code": "44",  "unicode": "127798"},
        {"f": "Webdings", "code": "45",  "unicode": "127895"},
        {"f": "Webdings", "code": "46",  "unicode": "128638"},
        {"f": "Webdings", "code": "47",  "unicode": "128636"},
        {"f": "Webdings", "code": "48",  "unicode": "128469"},
        {"f": "Webdings", "code": "49",  "unicode": "128470"},
        {"f": "Webdings", "code": "50",  "unicode": "128471"},
        {"f": "Webdings", "code": "51",  "unicode": "9204"},
        {"f": "Webdings", "code": "52",  "unicode": "9205"},
        {"f": "Webdings", "code": "53",  "unicode": "9206"},
        {"f": "Webdings", "code": "54",  "unicode": "9207"},
        {"f": "Webdings", "code": "55",  "unicode": "9194"},
        {"f": "Webdings", "code": "56",  "unicode": "9193"},
        {"f": "Webdings", "code": "57",  "unicode": "9198"},
        {"f": "Webdings", "code": "58",  "unicode": "9197"},
        {"f": "Webdings", "code": "59",  "unicode": "9208"},
        {"f": "Webdings", "code": "60",  "unicode": "9209"},
        {"f": "Webdings", "code": "61",  "unicode": "9210"},
        {"f": "Webdings", "code": "62",  "unicode": "128474"},
        {"f": "Webdings", "code": "63",  "unicode": "128499"},
        {"f": "Webdings", "code": "64",  "unicode": "128736"},
        {"f": "Webdings", "code": "65",  "unicode": "127959"},
        {"f": "Webdings", "code": "66",  "unicode": "127960"},
        {"f": "Webdings", "code": "67",  "unicode": "127961"},
        {"f": "Webdings", "code": "68",  "unicode": "127962"},
        {"f": "Webdings", "code": "69",  "unicode": "127964"},
        {"f": "Webdings", "code": "70",  "unicode": "127981"},
        {"f": "Webdings", "code": "71",  "unicode": "127963"},
        {"f": "Webdings", "code": "72",  "unicode": "127968"},
        {"f": "Webdings", "code": "73",  "unicode": "127958"},
        {"f": "Webdings", "code": "74",  "unicode": "127965"},
        {"f": "Webdings", "code": "75",  "unicode": "128739"},
        {"f": "Webdings", "code": "76",  "unicode": "128269"},
        {"f": "Webdings", "code": "77",  "unicode": "127956"},
        {"f": "Webdings", "code": "78",  "unicode": "128065"},
        {"f": "Webdings", "code": "79",  "unicode": "128066"},
        {"f": "Webdings", "code": "80",  "unicode": "127966"},
        {"f": "Webdings", "code": "81",  "unicode": "127957"},
        {"f": "Webdings", "code": "82",  "unicode": "128740"},
        {"f": "Webdings", "code": "83",  "unicode": "127967"},
        {"f": "Webdings", "code": "84",  "unicode": "128755"},
        {"f": "Webdings", "code": "85",  "unicode": "128364"},
        {"f": "Webdings", "code": "86",  "unicode": "128363"},
        {"f": "Webdings", "code": "87",  "unicode": "128360"},
        {"f": "Webdings", "code": "88",  "unicode": "128264"},
        {"f": "Webdings", "code": "89",  "unicode": "127892"},
        {"f": "Webdings", "code": "90",  "unicode": "127893"},
        {"f": "Webdings", "code": "91",  "unicode": "128492"},
        {"f": "Webdings", "code": "92",  "unicode": "128637"},
        {"f": "Webdings", "code": "93",  "unicode": "128493"},
        {"f": "Webdings", "code": "94",  "unicode": "128490"},
        {"f": "Webdings", "code": "95",  "unicode": "128491"},
        {"f": "Webdings", "code": "96",  "unicode": "11156"},
        {"f": "Webdings", "code": "97",  "unicode": "10004"},
        {"f": "Webdings", "code": "98",  "unicode": "128690"},
        {"f": "Webdings", "code": "99",  "unicode": "11036"},
        {"f": "Webdings", "code": "100",  "unicode": "128737"},
        {"f": "Webdings", "code": "101",  "unicode": "128230"},
        {"f": "Webdings", "code": "102",  "unicode": "128753"},
        {"f": "Webdings", "code": "103",  "unicode": "11035"},
        {"f": "Webdings", "code": "104",  "unicode": "128657"},
        {"f": "Webdings", "code": "105",  "unicode": "128712"},
        {"f": "Webdings", "code": "106",  "unicode": "128745"},
        {"f": "Webdings", "code": "107",  "unicode": "128752"},
        {"f": "Webdings", "code": "108",  "unicode": "128968"},
        {"f": "Webdings", "code": "109",  "unicode": "128372"},
        {"f": "Webdings", "code": "110",  "unicode": "11044"},
        {"f": "Webdings", "code": "111",  "unicode": "128741"},
        {"f": "Webdings", "code": "112",  "unicode": "128660"},
        {"f": "Webdings", "code": "113",  "unicode": "128472"},
        {"f": "Webdings", "code": "114",  "unicode": "128473"},
        {"f": "Webdings", "code": "115",  "unicode": "10067"},
        {"f": "Webdings", "code": "116",  "unicode": "128754"},
        {"f": "Webdings", "code": "117",  "unicode": "128647"},
        {"f": "Webdings", "code": "118",  "unicode": "128653"},
        {"f": "Webdings", "code": "119",  "unicode": "9971"},
        {"f": "Webdings", "code": "120",  "unicode": "10680"},
        {"f": "Webdings", "code": "121",  "unicode": "8854"},
        {"f": "Webdings", "code": "122",  "unicode": "128685"},
        {"f": "Webdings", "code": "123",  "unicode": "128494"},
        {"f": "Webdings", "code": "124",  "unicode": "9168"},
        {"f": "Webdings", "code": "125",  "unicode": "128495"},
        {"f": "Webdings", "code": "126",  "unicode": "128498"},
        {"f": "Webdings", "code": "128",  "unicode": "128697"},
        {"f": "Webdings", "code": "129",  "unicode": "128698"},
        {"f": "Webdings", "code": "130",  "unicode": "128713"},
        {"f": "Webdings", "code": "131",  "unicode": "128714"},
        {"f": "Webdings", "code": "132",  "unicode": "128700"},
        {"f": "Webdings", "code": "133",  "unicode": "128125"},
        {"f": "Webdings", "code": "134",  "unicode": "127947"},
        {"f": "Webdings", "code": "135",  "unicode": "9975"},
        {"f": "Webdings", "code": "136",  "unicode": "127938"},
        {"f": "Webdings", "code": "137",  "unicode": "127948"},
        {"f": "Webdings", "code": "138",  "unicode": "127946"},
        {"f": "Webdings", "code": "139",  "unicode": "127940"},
        {"f": "Webdings", "code": "140",  "unicode": "127949"},
        {"f": "Webdings", "code": "141",  "unicode": "127950"},
        {"f": "Webdings", "code": "142",  "unicode": "128664"},
        {"f": "Webdings", "code": "143",  "unicode": "128480"},
        {"f": "Webdings", "code": "144",  "unicode": "128738"},
        {"f": "Webdings", "code": "145",  "unicode": "128176"},
        {"f": "Webdings", "code": "146",  "unicode": "127991"},
        {"f": "Webdings", "code": "147",  "unicode": "128179"},
        {"f": "Webdings", "code": "148",  "unicode": "128106"},
        {"f": "Webdings", "code": "149",  "unicode": "128481"},
        {"f": "Webdings", "code": "150",  "unicode": "128482"},
        {"f": "Webdings", "code": "151",  "unicode": "128483"},
        {"f": "Webdings", "code": "152",  "unicode": "10031"},
        {"f": "Webdings", "code": "153",  "unicode": "128388"},
        {"f": "Webdings", "code": "154",  "unicode": "128389"},
        {"f": "Webdings", "code": "155",  "unicode": "128387"},
        {"f": "Webdings", "code": "156",  "unicode": "128390"},
        {"f": "Webdings", "code": "157",  "unicode": "128441"},
        {"f": "Webdings", "code": "158",  "unicode": "128442"},
        {"f": "Webdings", "code": "159",  "unicode": "128443"},
        {"f": "Webdings", "code": "160",  "unicode": "128373"},
        {"f": "Webdings", "code": "161",  "unicode": "128368"},
        {"f": "Webdings", "code": "162",  "unicode": "128445"},
        {"f": "Webdings", "code": "163",  "unicode": "128446"},
        {"f": "Webdings", "code": "164",  "unicode": "128203"},
        {"f": "Webdings", "code": "165",  "unicode": "128466"},
        {"f": "Webdings", "code": "166",  "unicode": "128467"},
        {"f": "Webdings", "code": "167",  "unicode": "128366"},
        {"f": "Webdings", "code": "168",  "unicode": "128218"},
        {"f": "Webdings", "code": "169",  "unicode": "128478"},
        {"f": "Webdings", "code": "170",  "unicode": "128479"},
        {"f": "Webdings", "code": "171",  "unicode": "128451"},
        {"f": "Webdings", "code": "172",  "unicode": "128450"},
        {"f": "Webdings", "code": "173",  "unicode": "128444"},
        {"f": "Webdings", "code": "174",  "unicode": "127917"},
        {"f": "Webdings", "code": "175",  "unicode": "127900"},
        {"f": "Webdings", "code": "176",  "unicode": "127896"},
        {"f": "Webdings", "code": "177",  "unicode": "127897"},
        {"f": "Webdings", "code": "178",  "unicode": "127911"},
        {"f": "Webdings", "code": "179",  "unicode": "128191"},
        {"f": "Webdings", "code": "180",  "unicode": "127902"},
        {"f": "Webdings", "code": "181",  "unicode": "128247"},
        {"f": "Webdings", "code": "182",  "unicode": "127903"},
        {"f": "Webdings", "code": "183",  "unicode": "127916"},
        {"f": "Webdings", "code": "184",  "unicode": "128253"},
        {"f": "Webdings", "code": "185",  "unicode": "128249"},
        {"f": "Webdings", "code": "186",  "unicode": "128254"},
        {"f": "Webdings", "code": "187",  "unicode": "128251"},
        {"f": "Webdings", "code": "188",  "unicode": "127898"},
        {"f": "Webdings", "code": "189",  "unicode": "127899"},
        {"f": "Webdings", "code": "190",  "unicode": "128250"},
        {"f": "Webdings", "code": "191",  "unicode": "128187"},
        {"f": "Webdings", "code": "192",  "unicode": "128421"},
        {"f": "Webdings", "code": "193",  "unicode": "128422"},
        {"f": "Webdings", "code": "194",  "unicode": "128423"},
        {"f": "Webdings", "code": "195",  "unicode": "128377"},
        {"f": "Webdings", "code": "196",  "unicode": "127918"},
        {"f": "Webdings", "code": "197",  "unicode": "128379"},
        {"f": "Webdings", "code": "198",  "unicode": "128380"},
        {"f": "Webdings", "code": "199",  "unicode": "128223"},
        {"f": "Webdings", "code": "200",  "unicode": "128385"},
        {"f": "Webdings", "code": "201",  "unicode": "128384"},
        {"f": "Webdings", "code": "202",  "unicode": "128424"},
        {"f": "Webdings", "code": "203",  "unicode": "128425"},
        {"f": "Webdings", "code": "204",  "unicode": "128447"},
        {"f": "Webdings", "code": "205",  "unicode": "128426"},
        {"f": "Webdings", "code": "206",  "unicode": "128476"},
        {"f": "Webdings", "code": "207",  "unicode": "128274"},
        {"f": "Webdings", "code": "208",  "unicode": "128275"},
        {"f": "Webdings", "code": "209",  "unicode": "128477"},
        {"f": "Webdings", "code": "210",  "unicode": "128229"},
        {"f": "Webdings", "code": "211",  "unicode": "128228"},
        {"f": "Webdings", "code": "212",  "unicode": "128371"},
        {"f": "Webdings", "code": "213",  "unicode": "127779"},
        {"f": "Webdings", "code": "214",  "unicode": "127780"},
        {"f": "Webdings", "code": "215",  "unicode": "127781"},
        {"f": "Webdings", "code": "216",  "unicode": "127782"},
        {"f": "Webdings", "code": "217",  "unicode": "9729"},
        {"f": "Webdings", "code": "218",  "unicode": "127784"},
        {"f": "Webdings", "code": "219",  "unicode": "127783"},
        {"f": "Webdings", "code": "220",  "unicode": "127785"},
        {"f": "Webdings", "code": "221",  "unicode": "127786"},
        {"f": "Webdings", "code": "222",  "unicode": "127788"},
        {"f": "Webdings", "code": "223",  "unicode": "127787"},
        {"f": "Webdings", "code": "224",  "unicode": "127772"},
        {"f": "Webdings", "code": "225",  "unicode": "127777"},
        {"f": "Webdings", "code": "226",  "unicode": "128715"},
        {"f": "Webdings", "code": "227",  "unicode": "128719"},
        {"f": "Webdings", "code": "228",  "unicode": "127869"},
        {"f": "Webdings", "code": "229",  "unicode": "127864"},
        {"f": "Webdings", "code": "230",  "unicode": "128718"},
        {"f": "Webdings", "code": "231",  "unicode": "128717"},
        {"f": "Webdings", "code": "232",  "unicode": "9413"},
        {"f": "Webdings", "code": "233",  "unicode": "9855"},
        {"f": "Webdings", "code": "234",  "unicode": "128710"},
        {"f": "Webdings", "code": "235",  "unicode": "128392"},
        {"f": "Webdings", "code": "236",  "unicode": "127891"},
        {"f": "Webdings", "code": "237",  "unicode": "128484"},
        {"f": "Webdings", "code": "238",  "unicode": "128485"},
        {"f": "Webdings", "code": "239",  "unicode": "128486"},
        {"f": "Webdings", "code": "240",  "unicode": "128487"},
        {"f": "Webdings", "code": "241",  "unicode": "128746"},
        {"f": "Webdings", "code": "242",  "unicode": "128063"},
        {"f": "Webdings", "code": "243",  "unicode": "128038"},
        {"f": "Webdings", "code": "244",  "unicode": "128031"},
        {"f": "Webdings", "code": "245",  "unicode": "128021"},
        {"f": "Webdings", "code": "246",  "unicode": "128008"},
        {"f": "Webdings", "code": "247",  "unicode": "128620"},
        {"f": "Webdings", "code": "248",  "unicode": "128622"},
        {"f": "Webdings", "code": "249",  "unicode": "128621"},
        {"f": "Webdings", "code": "250",  "unicode": "128623"},
        {"f": "Webdings", "code": "251",  "unicode": "128506"},
        {"f": "Webdings", "code": "252",  "unicode": "127757"},
        {"f": "Webdings", "code": "253",  "unicode": "127759"},
        {"f": "Webdings", "code": "254",  "unicode": "127758"},
        {"f": "Webdings", "code": "255",  "unicode": "128330"},
        {"f": "Wingdings", "code": "32",  "unicode": "32"},
        {"f": "Wingdings", "code": "33",  "unicode": "128393"},
        {"f": "Wingdings", "code": "34",  "unicode": "9986"},
        {"f": "Wingdings", "code": "35",  "unicode": "9985"},
        {"f": "Wingdings", "code": "36",  "unicode": "128083"},
        {"f": "Wingdings", "code": "37",  "unicode": "128365"},
        {"f": "Wingdings", "code": "38",  "unicode": "128366"},
        {"f": "Wingdings", "code": "39",  "unicode": "128367"},
        {"f": "Wingdings", "code": "40",  "unicode": "128383"},
        {"f": "Wingdings", "code": "41",  "unicode": "9990"},
        {"f": "Wingdings", "code": "42",  "unicode": "128386"},
        {"f": "Wingdings", "code": "43",  "unicode": "128387"},
        {"f": "Wingdings", "code": "44",  "unicode": "128234"},
        {"f": "Wingdings", "code": "45",  "unicode": "128235"},
        {"f": "Wingdings", "code": "46",  "unicode": "128236"},
        {"f": "Wingdings", "code": "47",  "unicode": "128237"},
        {"f": "Wingdings", "code": "48",  "unicode": "128448"},
        {"f": "Wingdings", "code": "49",  "unicode": "128449"},
        {"f": "Wingdings", "code": "50",  "unicode": "128462"},
        {"f": "Wingdings", "code": "51",  "unicode": "128463"},
        {"f": "Wingdings", "code": "52",  "unicode": "128464"},
        {"f": "Wingdings", "code": "53",  "unicode": "128452"},
        {"f": "Wingdings", "code": "54",  "unicode": "8987"},
        {"f": "Wingdings", "code": "55",  "unicode": "128430"},
        {"f": "Wingdings", "code": "56",  "unicode": "128432"},
        {"f": "Wingdings", "code": "57",  "unicode": "128434"},
        {"f": "Wingdings", "code": "58",  "unicode": "128435"},
        {"f": "Wingdings", "code": "59",  "unicode": "128436"},
        {"f": "Wingdings", "code": "60",  "unicode": "128427"},
        {"f": "Wingdings", "code": "61",  "unicode": "128428"},
        {"f": "Wingdings", "code": "62",  "unicode": "9991"},
        {"f": "Wingdings", "code": "63",  "unicode": "9997"},
        {"f": "Wingdings", "code": "64",  "unicode": "128398"},
        {"f": "Wingdings", "code": "65",  "unicode": "9996"},
        {"f": "Wingdings", "code": "66",  "unicode": "128399"},
        {"f": "Wingdings", "code": "67",  "unicode": "128077"},
        {"f": "Wingdings", "code": "68",  "unicode": "128078"},
        {"f": "Wingdings", "code": "69",  "unicode": "9756"},
        {"f": "Wingdings", "code": "70",  "unicode": "9758"},
        {"f": "Wingdings", "code": "71",  "unicode": "9757"},
        {"f": "Wingdings", "code": "72",  "unicode": "9759"},
        {"f": "Wingdings", "code": "73",  "unicode": "128400"},
        {"f": "Wingdings", "code": "74",  "unicode": "9786"},
        {"f": "Wingdings", "code": "75",  "unicode": "128528"},
        {"f": "Wingdings", "code": "76",  "unicode": "9785"},
        {"f": "Wingdings", "code": "77",  "unicode": "128163"},
        {"f": "Wingdings", "code": "78",  "unicode": "128369"},
        {"f": "Wingdings", "code": "79",  "unicode": "127987"},
        {"f": "Wingdings", "code": "80",  "unicode": "127985"},
        {"f": "Wingdings", "code": "81",  "unicode": "9992"},
        {"f": "Wingdings", "code": "82",  "unicode": "9788"},
        {"f": "Wingdings", "code": "83",  "unicode": "127778"},
        {"f": "Wingdings", "code": "84",  "unicode": "10052"},
        {"f": "Wingdings", "code": "85",  "unicode": "128326"},
        {"f": "Wingdings", "code": "86",  "unicode": "10014"},
        {"f": "Wingdings", "code": "87",  "unicode": "128328"},
        {"f": "Wingdings", "code": "88",  "unicode": "10016"},
        {"f": "Wingdings", "code": "89",  "unicode": "10017"},
        {"f": "Wingdings", "code": "90",  "unicode": "9770"},
        {"f": "Wingdings", "code": "91",  "unicode": "9775"},
        {"f": "Wingdings", "code": "92",  "unicode": "128329"},
        {"f": "Wingdings", "code": "93",  "unicode": "9784"},
        {"f": "Wingdings", "code": "94",  "unicode": "9800"},
        {"f": "Wingdings", "code": "95",  "unicode": "9801"},
        {"f": "Wingdings", "code": "96",  "unicode": "9802"},
        {"f": "Wingdings", "code": "97",  "unicode": "9803"},
        {"f": "Wingdings", "code": "98",  "unicode": "9804"},
        {"f": "Wingdings", "code": "99",  "unicode": "9805"},
        {"f": "Wingdings", "code": "100",  "unicode": "9806"},
        {"f": "Wingdings", "code": "101",  "unicode": "9807"},
        {"f": "Wingdings", "code": "102",  "unicode": "9808"},
        {"f": "Wingdings", "code": "103",  "unicode": "9809"},
        {"f": "Wingdings", "code": "104",  "unicode": "9810"},
        {"f": "Wingdings", "code": "105",  "unicode": "9811"},
        {"f": "Wingdings", "code": "106",  "unicode": "128624"},
        {"f": "Wingdings", "code": "107",  "unicode": "128629"},
        {"f": "Wingdings", "code": "108",  "unicode": "9899"},
        {"f": "Wingdings", "code": "109",  "unicode": "128318"},
        {"f": "Wingdings", "code": "110",  "unicode": "9724"},
        {"f": "Wingdings", "code": "111",  "unicode": "128911"},
        {"f": "Wingdings", "code": "112",  "unicode": "128912"},
        {"f": "Wingdings", "code": "113",  "unicode": "10065"},
        {"f": "Wingdings", "code": "114",  "unicode": "10066"},
        {"f": "Wingdings", "code": "115",  "unicode": "128927"},
        {"f": "Wingdings", "code": "116",  "unicode": "10731"},
        {"f": "Wingdings", "code": "117",  "unicode": "9670"},
        {"f": "Wingdings", "code": "118",  "unicode": "10070"},
        {"f": "Wingdings", "code": "119",  "unicode": "11049"},
        {"f": "Wingdings", "code": "120",  "unicode": "8999"},
        {"f": "Wingdings", "code": "121",  "unicode": "11193"},
        {"f": "Wingdings", "code": "122",  "unicode": "8984"},
        {"f": "Wingdings", "code": "123",  "unicode": "127989"},
        {"f": "Wingdings", "code": "124",  "unicode": "127990"},
        {"f": "Wingdings", "code": "125",  "unicode": "128630"},
        {"f": "Wingdings", "code": "126",  "unicode": "128631"},
        {"f": "Wingdings", "code": "127",  "unicode": "9647"},
        {"f": "Wingdings", "code": "128",  "unicode": "127243"},
        {"f": "Wingdings", "code": "129",  "unicode": "10112"},
        {"f": "Wingdings", "code": "130",  "unicode": "10113"},
        {"f": "Wingdings", "code": "131",  "unicode": "10114"},
        {"f": "Wingdings", "code": "132",  "unicode": "10115"},
        {"f": "Wingdings", "code": "133",  "unicode": "10116"},
        {"f": "Wingdings", "code": "134",  "unicode": "10117"},
        {"f": "Wingdings", "code": "135",  "unicode": "10118"},
        {"f": "Wingdings", "code": "136",  "unicode": "10119"},
        {"f": "Wingdings", "code": "137",  "unicode": "10120"},
        {"f": "Wingdings", "code": "138",  "unicode": "10121"},
        {"f": "Wingdings", "code": "139",  "unicode": "127244"},
        {"f": "Wingdings", "code": "140",  "unicode": "10122"},
        {"f": "Wingdings", "code": "141",  "unicode": "10123"},
        {"f": "Wingdings", "code": "142",  "unicode": "10124"},
        {"f": "Wingdings", "code": "143",  "unicode": "10125"},
        {"f": "Wingdings", "code": "144",  "unicode": "10126"},
        {"f": "Wingdings", "code": "145",  "unicode": "10127"},
        {"f": "Wingdings", "code": "146",  "unicode": "10128"},
        {"f": "Wingdings", "code": "147",  "unicode": "10129"},
        {"f": "Wingdings", "code": "148",  "unicode": "10130"},
        {"f": "Wingdings", "code": "149",  "unicode": "10131"},
        {"f": "Wingdings", "code": "150",  "unicode": "128610"},
        {"f": "Wingdings", "code": "151",  "unicode": "128608"},
        {"f": "Wingdings", "code": "152",  "unicode": "128609"},
        {"f": "Wingdings", "code": "153",  "unicode": "128611"},
        {"f": "Wingdings", "code": "154",  "unicode": "128606"},
        {"f": "Wingdings", "code": "155",  "unicode": "128604"},
        {"f": "Wingdings", "code": "156",  "unicode": "128605"},
        {"f": "Wingdings", "code": "157",  "unicode": "128607"},
        {"f": "Wingdings", "code": "158",  "unicode": "8729"},
        {"f": "Wingdings", "code": "159",  "unicode": "8226"},
        {"f": "Wingdings", "code": "160",  "unicode": "11037"},
        {"f": "Wingdings", "code": "161",  "unicode": "11096"},
        {"f": "Wingdings", "code": "162",  "unicode": "128902"},
        {"f": "Wingdings", "code": "163",  "unicode": "128904"},
        {"f": "Wingdings", "code": "164",  "unicode": "128906"},
        {"f": "Wingdings", "code": "165",  "unicode": "128907"},
        {"f": "Wingdings", "code": "166",  "unicode": "128319"},
        {"f": "Wingdings", "code": "167",  "unicode": "9642"},
        {"f": "Wingdings", "code": "168",  "unicode": "128910"},
        {"f": "Wingdings", "code": "169",  "unicode": "128961"},
        {"f": "Wingdings", "code": "170",  "unicode": "128965"},
        {"f": "Wingdings", "code": "171",  "unicode": "9733"},
        {"f": "Wingdings", "code": "172",  "unicode": "128971"},
        {"f": "Wingdings", "code": "173",  "unicode": "128975"},
        {"f": "Wingdings", "code": "174",  "unicode": "128979"},
        {"f": "Wingdings", "code": "175",  "unicode": "128977"},
        {"f": "Wingdings", "code": "176",  "unicode": "11216"},
        {"f": "Wingdings", "code": "177",  "unicode": "8982"},
        {"f": "Wingdings", "code": "178",  "unicode": "11214"},
        {"f": "Wingdings", "code": "179",  "unicode": "11215"},
        {"f": "Wingdings", "code": "180",  "unicode": "11217"},
        {"f": "Wingdings", "code": "181",  "unicode": "10026"},
        {"f": "Wingdings", "code": "182",  "unicode": "10032"},
        {"f": "Wingdings", "code": "183",  "unicode": "128336"},
        {"f": "Wingdings", "code": "184",  "unicode": "128337"},
        {"f": "Wingdings", "code": "185",  "unicode": "128338"},
        {"f": "Wingdings", "code": "186",  "unicode": "128339"},
        {"f": "Wingdings", "code": "187",  "unicode": "128340"},
        {"f": "Wingdings", "code": "188",  "unicode": "128341"},
        {"f": "Wingdings", "code": "189",  "unicode": "128342"},
        {"f": "Wingdings", "code": "190",  "unicode": "128343"},
        {"f": "Wingdings", "code": "191",  "unicode": "128344"},
        {"f": "Wingdings", "code": "192",  "unicode": "128345"},
        {"f": "Wingdings", "code": "193",  "unicode": "128346"},
        {"f": "Wingdings", "code": "194",  "unicode": "128347"},
        {"f": "Wingdings", "code": "195",  "unicode": "11184"},
        {"f": "Wingdings", "code": "196",  "unicode": "11185"},
        {"f": "Wingdings", "code": "197",  "unicode": "11186"},
        {"f": "Wingdings", "code": "198",  "unicode": "11187"},
        {"f": "Wingdings", "code": "199",  "unicode": "11188"},
        {"f": "Wingdings", "code": "200",  "unicode": "11189"},
        {"f": "Wingdings", "code": "201",  "unicode": "11190"},
        {"f": "Wingdings", "code": "202",  "unicode": "11191"},
        {"f": "Wingdings", "code": "203",  "unicode": "128618"},
        {"f": "Wingdings", "code": "204",  "unicode": "128619"},
        {"f": "Wingdings", "code": "205",  "unicode": "128597"},
        {"f": "Wingdings", "code": "206",  "unicode": "128596"},
        {"f": "Wingdings", "code": "207",  "unicode": "128599"},
        {"f": "Wingdings", "code": "208",  "unicode": "128598"},
        {"f": "Wingdings", "code": "209",  "unicode": "128592"},
        {"f": "Wingdings", "code": "210",  "unicode": "128593"},
        {"f": "Wingdings", "code": "211",  "unicode": "128594"},
        {"f": "Wingdings", "code": "212",  "unicode": "128595"},
        {"f": "Wingdings", "code": "213",  "unicode": "9003"},
        {"f": "Wingdings", "code": "214",  "unicode": "8998"},
        {"f": "Wingdings", "code": "215",  "unicode": "11160"},
        {"f": "Wingdings", "code": "216",  "unicode": "11162"},
        {"f": "Wingdings", "code": "217",  "unicode": "11161"},
        {"f": "Wingdings", "code": "218",  "unicode": "11163"},
        {"f": "Wingdings", "code": "219",  "unicode": "11144"},
        {"f": "Wingdings", "code": "220",  "unicode": "11146"},
        {"f": "Wingdings", "code": "221",  "unicode": "11145"},
        {"f": "Wingdings", "code": "222",  "unicode": "11147"},
        {"f": "Wingdings", "code": "223",  "unicode": "129128"},
        {"f": "Wingdings", "code": "224",  "unicode": "129130"},
        {"f": "Wingdings", "code": "225",  "unicode": "129129"},
        {"f": "Wingdings", "code": "226",  "unicode": "129131"},
        {"f": "Wingdings", "code": "227",  "unicode": "129132"},
        {"f": "Wingdings", "code": "228",  "unicode": "129133"},
        {"f": "Wingdings", "code": "229",  "unicode": "129135"},
        {"f": "Wingdings", "code": "230",  "unicode": "129134"},
        {"f": "Wingdings", "code": "231",  "unicode": "129144"},
        {"f": "Wingdings", "code": "232",  "unicode": "129146"},
        {"f": "Wingdings", "code": "233",  "unicode": "129145"},
        {"f": "Wingdings", "code": "234",  "unicode": "129147"},
        {"f": "Wingdings", "code": "235",  "unicode": "129148"},
        {"f": "Wingdings", "code": "236",  "unicode": "129149"},
        {"f": "Wingdings", "code": "237",  "unicode": "129151"},
        {"f": "Wingdings", "code": "238",  "unicode": "129150"},
        {"f": "Wingdings", "code": "239",  "unicode": "8678"},
        {"f": "Wingdings", "code": "240",  "unicode": "8680"},
        {"f": "Wingdings", "code": "241",  "unicode": "8679"},
        {"f": "Wingdings", "code": "242",  "unicode": "8681"},
        {"f": "Wingdings", "code": "243",  "unicode": "11012"},
        {"f": "Wingdings", "code": "244",  "unicode": "8691"},
        {"f": "Wingdings", "code": "245",  "unicode": "11009"},
        {"f": "Wingdings", "code": "246",  "unicode": "11008"},
        {"f": "Wingdings", "code": "247",  "unicode": "11011"},
        {"f": "Wingdings", "code": "248",  "unicode": "11010"},
        {"f": "Wingdings", "code": "249",  "unicode": "129196"},
        {"f": "Wingdings", "code": "250",  "unicode": "129197"},
        {"f": "Wingdings", "code": "251",  "unicode": "128502"},
        {"f": "Wingdings", "code": "252",  "unicode": "10003"},
        {"f": "Wingdings", "code": "253",  "unicode": "128503"},
        {"f": "Wingdings", "code": "254",  "unicode": "128505"},
        {"f": "Wingdings 2", "code": "32",  "unicode": "32"},
        {"f": "Wingdings 2", "code": "33",  "unicode": "128394"},
        {"f": "Wingdings 2", "code": "34",  "unicode": "128395"},
        {"f": "Wingdings 2", "code": "35",  "unicode": "128396"},
        {"f": "Wingdings 2", "code": "36",  "unicode": "128397"},
        {"f": "Wingdings 2", "code": "37",  "unicode": "9988"},
        {"f": "Wingdings 2", "code": "38",  "unicode": "9984"},
        {"f": "Wingdings 2", "code": "39",  "unicode": "128382"},
        {"f": "Wingdings 2", "code": "40",  "unicode": "128381"},
        {"f": "Wingdings 2", "code": "41",  "unicode": "128453"},
        {"f": "Wingdings 2", "code": "42",  "unicode": "128454"},
        {"f": "Wingdings 2", "code": "43",  "unicode": "128455"},
        {"f": "Wingdings 2", "code": "44",  "unicode": "128456"},
        {"f": "Wingdings 2", "code": "45",  "unicode": "128457"},
        {"f": "Wingdings 2", "code": "46",  "unicode": "128458"},
        {"f": "Wingdings 2", "code": "47",  "unicode": "128459"},
        {"f": "Wingdings 2", "code": "48",  "unicode": "128460"},
        {"f": "Wingdings 2", "code": "49",  "unicode": "128461"},
        {"f": "Wingdings 2", "code": "50",  "unicode": "128203"},
        {"f": "Wingdings 2", "code": "51",  "unicode": "128465"},
        {"f": "Wingdings 2", "code": "52",  "unicode": "128468"},
        {"f": "Wingdings 2", "code": "53",  "unicode": "128437"},
        {"f": "Wingdings 2", "code": "54",  "unicode": "128438"},
        {"f": "Wingdings 2", "code": "55",  "unicode": "128439"},
        {"f": "Wingdings 2", "code": "56",  "unicode": "128440"},
        {"f": "Wingdings 2", "code": "57",  "unicode": "128429"},
        {"f": "Wingdings 2", "code": "58",  "unicode": "128431"},
        {"f": "Wingdings 2", "code": "59",  "unicode": "128433"},
        {"f": "Wingdings 2", "code": "60",  "unicode": "128402"},
        {"f": "Wingdings 2", "code": "61",  "unicode": "128403"},
        {"f": "Wingdings 2", "code": "62",  "unicode": "128408"},
        {"f": "Wingdings 2", "code": "63",  "unicode": "128409"},
        {"f": "Wingdings 2", "code": "64",  "unicode": "128410"},
        {"f": "Wingdings 2", "code": "65",  "unicode": "128411"},
        {"f": "Wingdings 2", "code": "66",  "unicode": "128072"},
        {"f": "Wingdings 2", "code": "67",  "unicode": "128073"},
        {"f": "Wingdings 2", "code": "68",  "unicode": "128412"},
        {"f": "Wingdings 2", "code": "69",  "unicode": "128413"},
        {"f": "Wingdings 2", "code": "70",  "unicode": "128414"},
        {"f": "Wingdings 2", "code": "71",  "unicode": "128415"},
        {"f": "Wingdings 2", "code": "72",  "unicode": "128416"},
        {"f": "Wingdings 2", "code": "73",  "unicode": "128417"},
        {"f": "Wingdings 2", "code": "74",  "unicode": "128070"},
        {"f": "Wingdings 2", "code": "75",  "unicode": "128071"},
        {"f": "Wingdings 2", "code": "76",  "unicode": "128418"},
        {"f": "Wingdings 2", "code": "77",  "unicode": "128419"},
        {"f": "Wingdings 2", "code": "78",  "unicode": "128401"},
        {"f": "Wingdings 2", "code": "79",  "unicode": "128500"},
        {"f": "Wingdings 2", "code": "80",  "unicode": "128504"},
        {"f": "Wingdings 2", "code": "81",  "unicode": "128501"},
        {"f": "Wingdings 2", "code": "82",  "unicode": "9745"},
        {"f": "Wingdings 2", "code": "83",  "unicode": "11197"},
        {"f": "Wingdings 2", "code": "84",  "unicode": "9746"},
        {"f": "Wingdings 2", "code": "85",  "unicode": "11198"},
        {"f": "Wingdings 2", "code": "86",  "unicode": "11199"},
        {"f": "Wingdings 2", "code": "87",  "unicode": "128711"},
        {"f": "Wingdings 2", "code": "88",  "unicode": "10680"},
        {"f": "Wingdings 2", "code": "89",  "unicode": "128625"},
        {"f": "Wingdings 2", "code": "90",  "unicode": "128628"},
        {"f": "Wingdings 2", "code": "91",  "unicode": "128626"},
        {"f": "Wingdings 2", "code": "92",  "unicode": "128627"},
        {"f": "Wingdings 2", "code": "93",  "unicode": "8253"},
        {"f": "Wingdings 2", "code": "94",  "unicode": "128633"},
        {"f": "Wingdings 2", "code": "95",  "unicode": "128634"},
        {"f": "Wingdings 2", "code": "96",  "unicode": "128635"},
        {"f": "Wingdings 2", "code": "97",  "unicode": "128614"},
        {"f": "Wingdings 2", "code": "98",  "unicode": "128612"},
        {"f": "Wingdings 2", "code": "99",  "unicode": "128613"},
        {"f": "Wingdings 2", "code": "100",  "unicode": "128615"},
        {"f": "Wingdings 2", "code": "101",  "unicode": "128602"},
        {"f": "Wingdings 2", "code": "102",  "unicode": "128600"},
        {"f": "Wingdings 2", "code": "103",  "unicode": "128601"},
        {"f": "Wingdings 2", "code": "104",  "unicode": "128603"},
        {"f": "Wingdings 2", "code": "105",  "unicode": "9450"},
        {"f": "Wingdings 2", "code": "106",  "unicode": "9312"},
        {"f": "Wingdings 2", "code": "107",  "unicode": "9313"},
        {"f": "Wingdings 2", "code": "108",  "unicode": "9314"},
        {"f": "Wingdings 2", "code": "109",  "unicode": "9315"},
        {"f": "Wingdings 2", "code": "110",  "unicode": "9316"},
        {"f": "Wingdings 2", "code": "111",  "unicode": "9317"},
        {"f": "Wingdings 2", "code": "112",  "unicode": "9318"},
        {"f": "Wingdings 2", "code": "113",  "unicode": "9319"},
        {"f": "Wingdings 2", "code": "114",  "unicode": "9320"},
        {"f": "Wingdings 2", "code": "115",  "unicode": "9321"},
        {"f": "Wingdings 2", "code": "116",  "unicode": "9471"},
        {"f": "Wingdings 2", "code": "117",  "unicode": "10102"},
        {"f": "Wingdings 2", "code": "118",  "unicode": "10103"},
        {"f": "Wingdings 2", "code": "119",  "unicode": "10104"},
        {"f": "Wingdings 2", "code": "120",  "unicode": "10105"},
        {"f": "Wingdings 2", "code": "121",  "unicode": "10106"},
        {"f": "Wingdings 2", "code": "122",  "unicode": "10107"},
        {"f": "Wingdings 2", "code": "123",  "unicode": "10108"},
        {"f": "Wingdings 2", "code": "124",  "unicode": "10109"},
        {"f": "Wingdings 2", "code": "125",  "unicode": "10110"},
        {"f": "Wingdings 2", "code": "126",  "unicode": "10111"},
        {"f": "Wingdings 2", "code": "128",  "unicode": "9737"},
        {"f": "Wingdings 2", "code": "129",  "unicode": "127765"},
        {"f": "Wingdings 2", "code": "130",  "unicode": "9789"},
        {"f": "Wingdings 2", "code": "131",  "unicode": "9790"},
        {"f": "Wingdings 2", "code": "132",  "unicode": "11839"},
        {"f": "Wingdings 2", "code": "133",  "unicode": "10013"},
        {"f": "Wingdings 2", "code": "134",  "unicode": "128327"},
        {"f": "Wingdings 2", "code": "135",  "unicode": "128348"},
        {"f": "Wingdings 2", "code": "136",  "unicode": "128349"},
        {"f": "Wingdings 2", "code": "137",  "unicode": "128350"},
        {"f": "Wingdings 2", "code": "138",  "unicode": "128351"},
        {"f": "Wingdings 2", "code": "139",  "unicode": "128352"},
        {"f": "Wingdings 2", "code": "140",  "unicode": "128353"},
        {"f": "Wingdings 2", "code": "141",  "unicode": "128354"},
        {"f": "Wingdings 2", "code": "142",  "unicode": "128355"},
        {"f": "Wingdings 2", "code": "143",  "unicode": "128356"},
        {"f": "Wingdings 2", "code": "144",  "unicode": "128357"},
        {"f": "Wingdings 2", "code": "145",  "unicode": "128358"},
        {"f": "Wingdings 2", "code": "146",  "unicode": "128359"},
        {"f": "Wingdings 2", "code": "147",  "unicode": "128616"},
        {"f": "Wingdings 2", "code": "148",  "unicode": "128617"},
        {"f": "Wingdings 2", "code": "149",  "unicode": "8901"},
        {"f": "Wingdings 2", "code": "150",  "unicode": "128900"},
        {"f": "Wingdings 2", "code": "151",  "unicode": "10625"},
        {"f": "Wingdings 2", "code": "152",  "unicode": "9679"},
        {"f": "Wingdings 2", "code": "153",  "unicode": "9675"},
        {"f": "Wingdings 2", "code": "154",  "unicode": "128901"},
        {"f": "Wingdings 2", "code": "155",  "unicode": "128903"},
        {"f": "Wingdings 2", "code": "156",  "unicode": "128905"},
        {"f": "Wingdings 2", "code": "157",  "unicode": "8857"},
        {"f": "Wingdings 2", "code": "158",  "unicode": "10687"},
        {"f": "Wingdings 2", "code": "159",  "unicode": "128908"},
        {"f": "Wingdings 2", "code": "160",  "unicode": "128909"},
        {"f": "Wingdings 2", "code": "161",  "unicode": "9726"},
        {"f": "Wingdings 2", "code": "162",  "unicode": "9632"},
        {"f": "Wingdings 2", "code": "163",  "unicode": "9633"},
        {"f": "Wingdings 2", "code": "164",  "unicode": "128913"},
        {"f": "Wingdings 2", "code": "165",  "unicode": "128914"},
        {"f": "Wingdings 2", "code": "166",  "unicode": "128915"},
        {"f": "Wingdings 2", "code": "167",  "unicode": "128916"},
        {"f": "Wingdings 2", "code": "168",  "unicode": "9635"},
        {"f": "Wingdings 2", "code": "169",  "unicode": "128917"},
        {"f": "Wingdings 2", "code": "170",  "unicode": "128918"},
        {"f": "Wingdings 2", "code": "171",  "unicode": "128919"},
        {"f": "Wingdings 2", "code": "172",  "unicode": "128920"},
        {"f": "Wingdings 2", "code": "173",  "unicode": "11049"},
        {"f": "Wingdings 2", "code": "174",  "unicode": "11045"},
        {"f": "Wingdings 2", "code": "175",  "unicode": "9671"},
        {"f": "Wingdings 2", "code": "176",  "unicode": "128922"},
        {"f": "Wingdings 2", "code": "177",  "unicode": "9672"},
        {"f": "Wingdings 2", "code": "178",  "unicode": "128923"},
        {"f": "Wingdings 2", "code": "179",  "unicode": "128924"},
        {"f": "Wingdings 2", "code": "180",  "unicode": "128925"},
        {"f": "Wingdings 2", "code": "181",  "unicode": "128926"},
        {"f": "Wingdings 2", "code": "182",  "unicode": "11050"},
        {"f": "Wingdings 2", "code": "183",  "unicode": "11047"},
        {"f": "Wingdings 2", "code": "184",  "unicode": "9674"},
        {"f": "Wingdings 2", "code": "185",  "unicode": "128928"},
        {"f": "Wingdings 2", "code": "186",  "unicode": "9686"},
        {"f": "Wingdings 2", "code": "187",  "unicode": "9687"},
        {"f": "Wingdings 2", "code": "188",  "unicode": "11210"},
        {"f": "Wingdings 2", "code": "189",  "unicode": "11211"},
        {"f": "Wingdings 2", "code": "190",  "unicode": "11200"},
        {"f": "Wingdings 2", "code": "191",  "unicode": "11201"},
        {"f": "Wingdings 2", "code": "192",  "unicode": "11039"},
        {"f": "Wingdings 2", "code": "193",  "unicode": "11202"},
        {"f": "Wingdings 2", "code": "194",  "unicode": "11043"},
        {"f": "Wingdings 2", "code": "195",  "unicode": "11042"},
        {"f": "Wingdings 2", "code": "196",  "unicode": "11203"},
        {"f": "Wingdings 2", "code": "197",  "unicode": "11204"},
        {"f": "Wingdings 2", "code": "198",  "unicode": "128929"},
        {"f": "Wingdings 2", "code": "199",  "unicode": "128930"},
        {"f": "Wingdings 2", "code": "200",  "unicode": "128931"},
        {"f": "Wingdings 2", "code": "201",  "unicode": "128932"},
        {"f": "Wingdings 2", "code": "202",  "unicode": "128933"},
        {"f": "Wingdings 2", "code": "203",  "unicode": "128934"},
        {"f": "Wingdings 2", "code": "204",  "unicode": "128935"},
        {"f": "Wingdings 2", "code": "205",  "unicode": "128936"},
        {"f": "Wingdings 2", "code": "206",  "unicode": "128937"},
        {"f": "Wingdings 2", "code": "207",  "unicode": "128938"},
        {"f": "Wingdings 2", "code": "208",  "unicode": "128939"},
        {"f": "Wingdings 2", "code": "209",  "unicode": "128940"},
        {"f": "Wingdings 2", "code": "210",  "unicode": "128941"},
        {"f": "Wingdings 2", "code": "211",  "unicode": "128942"},
        {"f": "Wingdings 2", "code": "212",  "unicode": "128943"},
        {"f": "Wingdings 2", "code": "213",  "unicode": "128944"},
        {"f": "Wingdings 2", "code": "214",  "unicode": "128945"},
        {"f": "Wingdings 2", "code": "215",  "unicode": "128946"},
        {"f": "Wingdings 2", "code": "216",  "unicode": "128947"},
        {"f": "Wingdings 2", "code": "217",  "unicode": "128948"},
        {"f": "Wingdings 2", "code": "218",  "unicode": "128949"},
        {"f": "Wingdings 2", "code": "219",  "unicode": "128950"},
        {"f": "Wingdings 2", "code": "220",  "unicode": "128951"},
        {"f": "Wingdings 2", "code": "221",  "unicode": "128952"},
        {"f": "Wingdings 2", "code": "222",  "unicode": "128953"},
        {"f": "Wingdings 2", "code": "223",  "unicode": "128954"},
        {"f": "Wingdings 2", "code": "224",  "unicode": "128955"},
        {"f": "Wingdings 2", "code": "225",  "unicode": "128956"},
        {"f": "Wingdings 2", "code": "226",  "unicode": "128957"},
        {"f": "Wingdings 2", "code": "227",  "unicode": "128958"},
        {"f": "Wingdings 2", "code": "228",  "unicode": "128959"},
        {"f": "Wingdings 2", "code": "229",  "unicode": "128960"},
        {"f": "Wingdings 2", "code": "230",  "unicode": "128962"},
        {"f": "Wingdings 2", "code": "231",  "unicode": "128964"},
        {"f": "Wingdings 2", "code": "232",  "unicode": "128966"},
        {"f": "Wingdings 2", "code": "233",  "unicode": "128969"},
        {"f": "Wingdings 2", "code": "234",  "unicode": "128970"},
        {"f": "Wingdings 2", "code": "235",  "unicode": "10038"},
        {"f": "Wingdings 2", "code": "236",  "unicode": "128972"},
        {"f": "Wingdings 2", "code": "237",  "unicode": "128974"},
        {"f": "Wingdings 2", "code": "238",  "unicode": "128976"},
        {"f": "Wingdings 2", "code": "239",  "unicode": "128978"},
        {"f": "Wingdings 2", "code": "240",  "unicode": "10041"},
        {"f": "Wingdings 2", "code": "241",  "unicode": "128963"},
        {"f": "Wingdings 2", "code": "242",  "unicode": "128967"},
        {"f": "Wingdings 2", "code": "243",  "unicode": "10031"},
        {"f": "Wingdings 2", "code": "244",  "unicode": "128973"},
        {"f": "Wingdings 2", "code": "245",  "unicode": "128980"},
        {"f": "Wingdings 2", "code": "246",  "unicode": "11212"},
        {"f": "Wingdings 2", "code": "247",  "unicode": "11213"},
        {"f": "Wingdings 2", "code": "248",  "unicode": "8251"},
        {"f": "Wingdings 2", "code": "249",  "unicode": "8258"},
        {"f": "Wingdings 3", "code": "32",  "unicode": "32"},
        {"f": "Wingdings 3", "code": "33",  "unicode": "11104"},
        {"f": "Wingdings 3", "code": "34",  "unicode": "11106"},
        {"f": "Wingdings 3", "code": "35",  "unicode": "11105"},
        {"f": "Wingdings 3", "code": "36",  "unicode": "11107"},
        {"f": "Wingdings 3", "code": "37",  "unicode": "11110"},
        {"f": "Wingdings 3", "code": "38",  "unicode": "11111"},
        {"f": "Wingdings 3", "code": "39",  "unicode": "11113"},
        {"f": "Wingdings 3", "code": "40",  "unicode": "11112"},
        {"f": "Wingdings 3", "code": "41",  "unicode": "11120"},
        {"f": "Wingdings 3", "code": "42",  "unicode": "11122"},
        {"f": "Wingdings 3", "code": "43",  "unicode": "11121"},
        {"f": "Wingdings 3", "code": "44",  "unicode": "11123"},
        {"f": "Wingdings 3", "code": "45",  "unicode": "11126"},
        {"f": "Wingdings 3", "code": "46",  "unicode": "11128"},
        {"f": "Wingdings 3", "code": "47",  "unicode": "11131"},
        {"f": "Wingdings 3", "code": "48",  "unicode": "11133"},
        {"f": "Wingdings 3", "code": "49",  "unicode": "11108"},
        {"f": "Wingdings 3", "code": "50",  "unicode": "11109"},
        {"f": "Wingdings 3", "code": "51",  "unicode": "11114"},
        {"f": "Wingdings 3", "code": "52",  "unicode": "11116"},
        {"f": "Wingdings 3", "code": "53",  "unicode": "11115"},
        {"f": "Wingdings 3", "code": "54",  "unicode": "11117"},
        {"f": "Wingdings 3", "code": "55",  "unicode": "11085"},
        {"f": "Wingdings 3", "code": "56",  "unicode": "11168"},
        {"f": "Wingdings 3", "code": "57",  "unicode": "11169"},
        {"f": "Wingdings 3", "code": "58",  "unicode": "11170"},
        {"f": "Wingdings 3", "code": "59",  "unicode": "11171"},
        {"f": "Wingdings 3", "code": "60",  "unicode": "11172"},
        {"f": "Wingdings 3", "code": "61",  "unicode": "11173"},
        {"f": "Wingdings 3", "code": "62",  "unicode": "11174"},
        {"f": "Wingdings 3", "code": "63",  "unicode": "11175"},
        {"f": "Wingdings 3", "code": "64",  "unicode": "11152"},
        {"f": "Wingdings 3", "code": "65",  "unicode": "11153"},
        {"f": "Wingdings 3", "code": "66",  "unicode": "11154"},
        {"f": "Wingdings 3", "code": "67",  "unicode": "11155"},
        {"f": "Wingdings 3", "code": "68",  "unicode": "11136"},
        {"f": "Wingdings 3", "code": "69",  "unicode": "11139"},
        {"f": "Wingdings 3", "code": "70",  "unicode": "11134"},
        {"f": "Wingdings 3", "code": "71",  "unicode": "11135"},
        {"f": "Wingdings 3", "code": "72",  "unicode": "11140"},
        {"f": "Wingdings 3", "code": "73",  "unicode": "11142"},
        {"f": "Wingdings 3", "code": "74",  "unicode": "11141"},
        {"f": "Wingdings 3", "code": "75",  "unicode": "11143"},
        {"f": "Wingdings 3", "code": "76",  "unicode": "11151"},
        {"f": "Wingdings 3", "code": "77",  "unicode": "11149"},
        {"f": "Wingdings 3", "code": "78",  "unicode": "11150"},
        {"f": "Wingdings 3", "code": "79",  "unicode": "11148"},
        {"f": "Wingdings 3", "code": "80",  "unicode": "11118"},
        {"f": "Wingdings 3", "code": "81",  "unicode": "11119"},
        {"f": "Wingdings 3", "code": "82",  "unicode": "9099"},
        {"f": "Wingdings 3", "code": "83",  "unicode": "8996"},
        {"f": "Wingdings 3", "code": "84",  "unicode": "8963"},
        {"f": "Wingdings 3", "code": "85",  "unicode": "8997"},
        {"f": "Wingdings 3", "code": "86",  "unicode": "9251"},
        {"f": "Wingdings 3", "code": "87",  "unicode": "9085"},
        {"f": "Wingdings 3", "code": "88",  "unicode": "8682"},
        {"f": "Wingdings 3", "code": "89",  "unicode": "11192"},
        {"f": "Wingdings 3", "code": "90",  "unicode": "129184"},
        {"f": "Wingdings 3", "code": "91",  "unicode": "129185"},
        {"f": "Wingdings 3", "code": "92",  "unicode": "129186"},
        {"f": "Wingdings 3", "code": "93",  "unicode": "129187"},
        {"f": "Wingdings 3", "code": "94",  "unicode": "129188"},
        {"f": "Wingdings 3", "code": "95",  "unicode": "129189"},
        {"f": "Wingdings 3", "code": "96",  "unicode": "129190"},
        {"f": "Wingdings 3", "code": "97",  "unicode": "129191"},
        {"f": "Wingdings 3", "code": "98",  "unicode": "129192"},
        {"f": "Wingdings 3", "code": "99",  "unicode": "129193"},
        {"f": "Wingdings 3", "code": "100",  "unicode": "129194"},
        {"f": "Wingdings 3", "code": "101",  "unicode": "129195"},
        {"f": "Wingdings 3", "code": "102",  "unicode": "129104"},
        {"f": "Wingdings 3", "code": "103",  "unicode": "129106"},
        {"f": "Wingdings 3", "code": "104",  "unicode": "129105"},
        {"f": "Wingdings 3", "code": "105",  "unicode": "129107"},
        {"f": "Wingdings 3", "code": "106",  "unicode": "129108"},
        {"f": "Wingdings 3", "code": "107",  "unicode": "129109"},
        {"f": "Wingdings 3", "code": "108",  "unicode": "129111"},
        {"f": "Wingdings 3", "code": "109",  "unicode": "129110"},
        {"f": "Wingdings 3", "code": "110",  "unicode": "129112"},
        {"f": "Wingdings 3", "code": "111",  "unicode": "129113"},
        {"f": "Wingdings 3", "code": "112",  "unicode": "9650"},
        {"f": "Wingdings 3", "code": "113",  "unicode": "9660"},
        {"f": "Wingdings 3", "code": "114",  "unicode": "9651"},
        {"f": "Wingdings 3", "code": "115",  "unicode": "9661"},
        {"f": "Wingdings 3", "code": "116",  "unicode": "9664"},
        {"f": "Wingdings 3", "code": "117",  "unicode": "9654"},
        {"f": "Wingdings 3", "code": "118",  "unicode": "9665"},
        {"f": "Wingdings 3", "code": "119",  "unicode": "9655"},
        {"f": "Wingdings 3", "code": "120",  "unicode": "9699"},
        {"f": "Wingdings 3", "code": "121",  "unicode": "9698"},
        {"f": "Wingdings 3", "code": "122",  "unicode": "9700"},
        {"f": "Wingdings 3", "code": "123",  "unicode": "9701"},
        {"f": "Wingdings 3", "code": "124",  "unicode": "128896"},
        {"f": "Wingdings 3", "code": "125",  "unicode": "128898"},
        {"f": "Wingdings 3", "code": "126",  "unicode": "128897"},
        {"f": "Wingdings 3", "code": "128",  "unicode": "128899"},
        {"f": "Wingdings 3", "code": "129",  "unicode": "11205"},
        {"f": "Wingdings 3", "code": "130",  "unicode": "11206"},
        {"f": "Wingdings 3", "code": "131",  "unicode": "11207"},
        {"f": "Wingdings 3", "code": "132",  "unicode": "11208"},
        {"f": "Wingdings 3", "code": "133",  "unicode": "11164"},
        {"f": "Wingdings 3", "code": "134",  "unicode": "11166"},
        {"f": "Wingdings 3", "code": "135",  "unicode": "11165"},
        {"f": "Wingdings 3", "code": "136",  "unicode": "11167"},
        {"f": "Wingdings 3", "code": "137",  "unicode": "129040"},
        {"f": "Wingdings 3", "code": "138",  "unicode": "129042"},
        {"f": "Wingdings 3", "code": "139",  "unicode": "129041"},
        {"f": "Wingdings 3", "code": "140",  "unicode": "129043"},
        {"f": "Wingdings 3", "code": "141",  "unicode": "129044"},
        {"f": "Wingdings 3", "code": "142",  "unicode": "129046"},
        {"f": "Wingdings 3", "code": "143",  "unicode": "129045"},
        {"f": "Wingdings 3", "code": "144",  "unicode": "129047"},
        {"f": "Wingdings 3", "code": "145",  "unicode": "129048"},
        {"f": "Wingdings 3", "code": "146",  "unicode": "129050"},
        {"f": "Wingdings 3", "code": "147",  "unicode": "129049"},
        {"f": "Wingdings 3", "code": "148",  "unicode": "129051"},
        {"f": "Wingdings 3", "code": "149",  "unicode": "129052"},
        {"f": "Wingdings 3", "code": "150",  "unicode": "129054"},
        {"f": "Wingdings 3", "code": "151",  "unicode": "129053"},
        {"f": "Wingdings 3", "code": "152",  "unicode": "129055"},
        {"f": "Wingdings 3", "code": "153",  "unicode": "129024"},
        {"f": "Wingdings 3", "code": "154",  "unicode": "129026"},
        {"f": "Wingdings 3", "code": "155",  "unicode": "129025"},
        {"f": "Wingdings 3", "code": "156",  "unicode": "129027"},
        {"f": "Wingdings 3", "code": "157",  "unicode": "129028"},
        {"f": "Wingdings 3", "code": "158",  "unicode": "129030"},
        {"f": "Wingdings 3", "code": "159",  "unicode": "129029"},
        {"f": "Wingdings 3", "code": "160",  "unicode": "129031"},
        {"f": "Wingdings 3", "code": "161",  "unicode": "129032"},
        {"f": "Wingdings 3", "code": "162",  "unicode": "129034"},
        {"f": "Wingdings 3", "code": "163",  "unicode": "129033"},
        {"f": "Wingdings 3", "code": "164",  "unicode": "129035"},
        {"f": "Wingdings 3", "code": "165",  "unicode": "129056"},
        {"f": "Wingdings 3", "code": "166",  "unicode": "129058"},
        {"f": "Wingdings 3", "code": "167",  "unicode": "129060"},
        {"f": "Wingdings 3", "code": "168",  "unicode": "129062"},
        {"f": "Wingdings 3", "code": "169",  "unicode": "129064"},
        {"f": "Wingdings 3", "code": "170",  "unicode": "129066"},
        {"f": "Wingdings 3", "code": "171",  "unicode": "129068"},
        {"f": "Wingdings 3", "code": "172",  "unicode": "129180"},
        {"f": "Wingdings 3", "code": "173",  "unicode": "129181"},
        {"f": "Wingdings 3", "code": "174",  "unicode": "129182"},
        {"f": "Wingdings 3", "code": "175",  "unicode": "129183"},
        {"f": "Wingdings 3", "code": "176",  "unicode": "129070"},
        {"f": "Wingdings 3", "code": "177",  "unicode": "129072"},
        {"f": "Wingdings 3", "code": "178",  "unicode": "129074"},
        {"f": "Wingdings 3", "code": "179",  "unicode": "129076"},
        {"f": "Wingdings 3", "code": "180",  "unicode": "129078"},
        {"f": "Wingdings 3", "code": "181",  "unicode": "129080"},
        {"f": "Wingdings 3", "code": "182",  "unicode": "129082"},
        {"f": "Wingdings 3", "code": "183",  "unicode": "129081"},
        {"f": "Wingdings 3", "code": "184",  "unicode": "129083"},
        {"f": "Wingdings 3", "code": "185",  "unicode": "129176"},
        {"f": "Wingdings 3", "code": "186",  "unicode": "129178"},
        {"f": "Wingdings 3", "code": "187",  "unicode": "129177"},
        {"f": "Wingdings 3", "code": "188",  "unicode": "129179"},
        {"f": "Wingdings 3", "code": "189",  "unicode": "129084"},
        {"f": "Wingdings 3", "code": "190",  "unicode": "129086"},
        {"f": "Wingdings 3", "code": "191",  "unicode": "129085"},
        {"f": "Wingdings 3", "code": "192",  "unicode": "129087"},
        {"f": "Wingdings 3", "code": "193",  "unicode": "129088"},
        {"f": "Wingdings 3", "code": "194",  "unicode": "129090"},
        {"f": "Wingdings 3", "code": "195",  "unicode": "129089"},
        {"f": "Wingdings 3", "code": "196",  "unicode": "129091"},
        {"f": "Wingdings 3", "code": "197",  "unicode": "129092"},
        {"f": "Wingdings 3", "code": "198",  "unicode": "129094"},
        {"f": "Wingdings 3", "code": "199",  "unicode": "129093"},
        {"f": "Wingdings 3", "code": "200",  "unicode": "129095"},
        {"f": "Wingdings 3", "code": "201",  "unicode": "11176"},
        {"f": "Wingdings 3", "code": "202",  "unicode": "11177"},
        {"f": "Wingdings 3", "code": "203",  "unicode": "11178"},
        {"f": "Wingdings 3", "code": "204",  "unicode": "11179"},
        {"f": "Wingdings 3", "code": "205",  "unicode": "11180"},
        {"f": "Wingdings 3", "code": "206",  "unicode": "11181"},
        {"f": "Wingdings 3", "code": "207",  "unicode": "11182"},
        {"f": "Wingdings 3", "code": "208",  "unicode": "11183"},
        {"f": "Wingdings 3", "code": "209",  "unicode": "129120"},
        {"f": "Wingdings 3", "code": "210",  "unicode": "129122"},
        {"f": "Wingdings 3", "code": "211",  "unicode": "129121"},
        {"f": "Wingdings 3", "code": "212",  "unicode": "129123"},
        {"f": "Wingdings 3", "code": "213",  "unicode": "129124"},
        {"f": "Wingdings 3", "code": "214",  "unicode": "129125"},
        {"f": "Wingdings 3", "code": "215",  "unicode": "129127"},
        {"f": "Wingdings 3", "code": "216",  "unicode": "129126"},
        {"f": "Wingdings 3", "code": "217",  "unicode": "129136"},
        {"f": "Wingdings 3", "code": "218",  "unicode": "129138"},
        {"f": "Wingdings 3", "code": "219",  "unicode": "129137"},
        {"f": "Wingdings 3", "code": "220",  "unicode": "129139"},
        {"f": "Wingdings 3", "code": "221",  "unicode": "129140"},
        {"f": "Wingdings 3", "code": "222",  "unicode": "129141"},
        {"f": "Wingdings 3", "code": "223",  "unicode": "129143"},
        {"f": "Wingdings 3", "code": "224",  "unicode": "129142"},
        {"f": "Wingdings 3", "code": "225",  "unicode": "129152"},
        {"f": "Wingdings 3", "code": "226",  "unicode": "129154"},
        {"f": "Wingdings 3", "code": "227",  "unicode": "129153"},
        {"f": "Wingdings 3", "code": "228",  "unicode": "129155"},
        {"f": "Wingdings 3", "code": "229",  "unicode": "129156"},
        {"f": "Wingdings 3", "code": "230",  "unicode": "129157"},
        {"f": "Wingdings 3", "code": "231",  "unicode": "129159"},
        {"f": "Wingdings 3", "code": "232",  "unicode": "129158"},
        {"f": "Wingdings 3", "code": "233",  "unicode": "129168"},
        {"f": "Wingdings 3", "code": "234",  "unicode": "129170"},
        {"f": "Wingdings 3", "code": "235",  "unicode": "129169"},
        {"f": "Wingdings 3", "code": "236",  "unicode": "129171"},
        {"f": "Wingdings 3", "code": "237",  "unicode": "129172"},
        {"f": "Wingdings 3", "code": "238",  "unicode": "129174"},
        {"f": "Wingdings 3", "code": "239",  "unicode": "129173"},
        {"f": "Wingdings 3", "code": "240",  "unicode": "129175"}
    ];


    function getTextWidth(html) {
        var div = document.createElement('div');
        div.style.position = 'absolute';
        div.style.float = 'left';
        div.style.whiteSpace = 'nowrap';
        div.style.visibility = 'hidden';
        div.innerHTML = html;
        document.body.appendChild(div);
        var width = div.offsetWidth;
        document.body.removeChild(div);
        return width;
    }

    function genTextBody(textBodyNode, spNode, slideLayoutSpNode, slideMasterSpNode, type, idx, warpObj, tbl_col_width) {
            var text = "";
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            if (textBodyNode === undefined) {
                return text;
            }
            //rtl : <p:txBody>
            //          <a:bodyPr wrap="square" rtlCol="1">

            var pFontStyle = PPTXXmlUtils.getTextByPathList(spNode, ["p:style", "a:fontRef"]);
            //console.log("genTextBody spNode: ", PPTXXmlUtils.getTextByPathList(spNode,["p:spPr","a:xfrm","a:ext"]));

            //var lstStyle = textBodyNode["a:lstStyle"];
            
            var apNode = textBodyNode["a:p"];
            if (apNode.constructor !== Array) {
                apNode = [apNode];
            }

            for (var i = 0; i < apNode.length; i++) {
                var pNode = apNode[i];
                var rNode = pNode["a:r"];
                var fldNode = pNode["a:fld"];
                var brNode = pNode["a:br"];
                if (rNode !== undefined) {
                    rNode = (rNode.constructor === Array) ? rNode : [rNode];
                }
                if (rNode !== undefined && fldNode !== undefined) {
                    fldNode = (fldNode.constructor === Array) ? fldNode : [fldNode];
                    rNode = rNode.concat(fldNode)
                }
                if (rNode !== undefined && brNode !== undefined) {
                    is_first_br = true;
                    brNode = (brNode.constructor === Array) ? brNode : [brNode];
                    brNode.forEach(function (item, indx) {
                        item.type = "br";
                    });
                    if (brNode.length > 1) {
                        brNode.shift();
                    }
                    rNode = rNode.concat(brNode)
                    //console.log("single a:p  rNode:", rNode, "brNode:", brNode )
                    rNode.sort(function (a, b) {
                        return a.attrs.order - b.attrs.order;
                    });
                    //console.log("sorted rNode:",rNode)
                }
                //rtlStr = "";//"dir='"+isRTL+"'";
                var styleText = "";
                var marginsVer = PPTXStyleUtils.getVerticalMargins(pNode, textBodyNode, type, idx, warpObj);
                if (marginsVer != "") {
                    styleText = marginsVer;
                }
                if (type == "body" || type == "obj" || type == "shape") {
                    styleText += "font-size: 0px;";
                    //styleText += "line-height: 0;";
                    styleText += "font-weight: 100;";
                    styleText += "font-style: normal;";
                }
                var cssName = "";

                if (styleText in warpObj.styleTable) {
                    cssName = warpObj.styleTable[styleText]["name"];
                } else {
                    cssName = "_css_" + (Object.keys(warpObj.styleTable).length + 1);
                    warpObj.styleTable[styleText] = {
                        "name": cssName,
                        "text": styleText
                    };
                }
                //console.log("textBodyNode: ", textBodyNode["a:lstStyle"])
                var prg_width_node = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cx"]);
                var prg_height_node;// = PPTXXmlUtils.getTextByPathList(spNode, ["p:spPr", "a:xfrm", "a:ext", "attrs", "cy"]);
                var sld_prg_width = ((prg_width_node !== undefined) ? ("width:" + (parseInt(prg_width_node) * slideFactor) + "px;") : "width:inherit;");
                var sld_prg_height = ((prg_height_node !== undefined) ? ("height:" + (parseInt(prg_height_node) * slideFactor) + "px;") : "");
                var prg_dir = PPTXStyleUtils.getPregraphDir(pNode, textBodyNode, idx, type, warpObj);
                text += "<div style='display: flex;" + sld_prg_width + sld_prg_height + "' class='slide-prgrph " + PPTXStyleUtils.getHorizontalAlign(pNode, textBodyNode, idx, type, prg_dir, warpObj) + " " +
                    prg_dir + " " + cssName + "' >";
                var buText_ary = genBuChar(pNode, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj);
                var isBullate = (buText_ary[0] !== undefined && buText_ary[0] !== null && buText_ary[0] != "" ) ? true : false;
                var bu_width = (buText_ary[1] !== undefined && buText_ary[1] !== null && isBullate) ? buText_ary[1] + buText_ary[2] : 0;
                text += (buText_ary[0] !== undefined) ? buText_ary[0]:"";
                //get text margin 
                var margin_ary = PPTXStyleUtils.getPregraphMargn(pNode, idx, type, isBullate, warpObj);
                var margin = margin_ary[0];
                var mrgin_val = margin_ary[1];
                if (prg_width_node === undefined && tbl_col_width !== undefined && prg_width_node != 0){
                    //sorce : table text
                    prg_width_node = tbl_col_width;
                }

                var prgrph_text = "";
                //var prgr_txt_art = [];
                var total_text_len = 0;
                if (rNode === undefined && pNode !== undefined) {
                    // without r
                    var prgr_text = genSpanElement(pNode, undefined, spNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, 1, warpObj, isBullate);
                    if (isBullate) {
                        total_text_len += getTextWidth(prgr_text);
                    }
                    prgrph_text += prgr_text;
                } else if (rNode !== undefined) {
                    // with multi r
                    for (var j = 0; j < rNode.length; j++) {
                        var prgr_text = genSpanElement(rNode[j], j, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNode.length, warpObj, isBullate);
                        if (isBullate) {
                            total_text_len += getTextWidth(prgr_text);
                        }
                        prgrph_text += prgr_text;
                    }
                }

                prg_width_node = parseInt(prg_width_node) * slideFactor - bu_width - mrgin_val;
                if (isBullate) {
                    //get prg_width_node if there is a bulltes
                    //console.log("total_text_len: ", total_text_len, "prg_width_node:", prg_width_node)

                    if (total_text_len < prg_width_node ){
                        prg_width_node = total_text_len + bu_width;
                    }
                }
                var prg_width = ((prg_width_node !== undefined) ? ("width:" + (prg_width_node )) + "px;" : "width:inherit;");
                text += "<div style='height: 100%;direction: initial;overflow-wrap:break-word;word-wrap: break-word;" + prg_width + margin + "' >";
                text += prgrph_text;
                text += "</div>";
                text += "</div>";
            }

            return text;
        }
        
        function genBuChar(node, i, spNode, textBodyNode, pFontStyle, idx, type, warpObj) {
            //console.log("genBuChar node: ", node, ", spNode: ", spNode, ", pFontStyle: ", pFontStyle, "type", type)
            ///////////////////////////////////////Amir///////////////////////////////
            var sldMstrTxtStyles = warpObj["slideMasterTextStyles"];
            var lstStyle = textBodyNode["a:lstStyle"];

            var rNode = PPTXXmlUtils.getTextByPathList(node, ["a:r"]);
            if (rNode !== undefined && rNode.constructor === Array) {
                rNode = rNode[0]; //bullet only to first "a:r"
            }
            var lvl = parseInt (PPTXXmlUtils.getTextByPathList(node["a:pPr"], ["attrs", "lvl"])) + 1;
            if (isNaN(lvl)) {
                lvl = 1;
            }
            var lvlStr = "a:lvl" + lvl + "pPr";
            var dfltBultColor, dfltBultSize, bultColor, bultSize, color_tye;

            if (rNode !== undefined) {
                dfltBultColor = PPTXStyleUtils.getFontColorPr(rNode, spNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
                color_tye = dfltBultColor[2];
                dfltBultSize = PPTXStyleUtils.getFontSize(rNode, textBodyNode, pFontStyle, lvl, type, warpObj);
            } else {
                return "";
            }
            //console.log("Bullet Size: " + bultSize);

            var bullet = "", marRStr = "", marLStr = "", margin_val=0, font_val=0;
            /////////////////////////////////////////////////////////////////


            var pPrNode = node["a:pPr"];
            var BullNONE = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buNone"]);
            if (BullNONE !== undefined) {
                return "";
            }

            var buType = "TYPE_NONE";

            var layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(node, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;

            var buChar = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buChar", "attrs", "char"]);
            var buNum = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buAutoNum", "attrs", "type"]);
            var buPic = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buBlip"]);
            if (buChar !== undefined) {
                buType = "TYPE_BULLET";
            }
            if (buNum !== undefined) {
                buType = "TYPE_NUMERIC";
            }
            if (buPic !== undefined) {
                buType = "TYPE_BULPIC";
            }

            var buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPts", "attrs", "val"]);
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buSzPct", "attrs", "val"]);
                if (buFontSize !== undefined) {
                    var prcnt = parseInt(buFontSize) / 100000;
                    //dfltBultSize = XXpt
                    //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                    var dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                    bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                }
            } else {
                bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
            }

            //get definde bullet COLOR
            var buClrNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buClr"]);


            if (buChar === undefined && buNum === undefined && buPic === undefined) {

                if (lstStyle !== undefined) {
                    BullNONE = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buNone"]);
                    if (BullNONE !== undefined) {
                        return "";
                    }
                    buType = "TYPE_NONE";
                    buChar = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buChar", "attrs", "char"]);
                    buNum = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buAutoNum", "attrs", "type"]);
                    buPic = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr,"a:buBlip"]);
                    if (buChar !== undefined) {
                        buType = "TYPE_BULLET";
                    }
                    if (buNum !== undefined) {
                        buType = "TYPE_NUMERIC";
                    }
                    if (buPic !== undefined) {
                        buType = "TYPE_BULPIC";
                    }
                    if (buChar !== undefined || buNum !== undefined || buPic !== undefined) {
                        pPrNode = lstStyle[lvlStr];
                    }
                }
            }
            if (buChar === undefined && buNum === undefined && buPic === undefined) {
                //check in slidelayout and masterlayout - TODO
                if (pPrNodeLaout !== undefined) {
                    BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buNone"]);
                    if (BullNONE !== undefined) {
                        return "";
                    }
                    buType = "TYPE_NONE";
                    buChar = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buChar", "attrs", "char"]);
                    buNum = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buAutoNum", "attrs", "type"]);
                    buPic = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buBlip"]);
                    if (buChar !== undefined) {
                        buType = "TYPE_BULLET";
                    }
                    if (buNum !== undefined) {
                        buType = "TYPE_NUMERIC";
                    }
                    if (buPic !== undefined) {
                        buType = "TYPE_BULPIC";
                    }
                }
                if (buChar === undefined && buNum === undefined && buPic === undefined) {
                    //masterlayout

                    if (pPrNodeMaster !== undefined) {
                        BullNONE = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buNone"]);
                        if (BullNONE !== undefined) {
                            return "";
                        }
                        buType = "TYPE_NONE";
                        buChar = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buChar", "attrs", "char"]);
                        buNum = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buAutoNum", "attrs", "type"]);
                        buPic = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buBlip"]);
                        if (buChar !== undefined) {
                            buType = "TYPE_BULLET";
                        }
                        if (buNum !== undefined) {
                            buType = "TYPE_NUMERIC";
                        }
                        if (buPic !== undefined) {
                            buType = "TYPE_BULPIC";
                        }
                    }

                }

            }
            //rtl
            var getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }
            var isRTL = false;
            if (getRtlVal !== undefined && getRtlVal == "1") {
                isRTL = true;
            }
            //align
            var alignNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "algn"]); //"l" | "ctr" | "r" | "just" | "justLow" | "dist" | "thaiDist
            if (alignNode === undefined) {
                alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "algn"]);
                if (alignNode === undefined) {
                    alignNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "algn"]);
                }
            }
            //indent?
            var indentNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "indent"]);
            if (indentNode === undefined) {
                indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "indent"]);
                if (indentNode === undefined) {
                    indentNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "indent"]);
                }
            }
            var indent = 0;
            if (indentNode !== undefined) {
                indent = parseInt(indentNode) * slideFactor;
            }
            //marL
            var marLNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"]);
            if (marLNode === undefined) {
                marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marL"]);
                if (marLNode === undefined) {
                    marLNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marL"]);
                }
            }
            //console.log("genBuChar() isRTL", isRTL, "alignNode:", alignNode)
            if (marLNode !== undefined) {
                var marginLeft = parseInt(marLNode) * slideFactor;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                margin_val = ((marginLeft + indent < 0) ? 0 : (marginLeft + indent));
                marLStr += margin_val + "px;";
            }
            
            //marR?
            var marRNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"]);
            if (marRNode === undefined && marLNode === undefined) {
                //need to check if this posble - TODO
                marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "marR"]);
                if (marRNode === undefined) {
                    marRNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "marR"]);
                }
            }
            if (marRNode !== undefined) {
                var marginRight = parseInt(marRNode) * slideFactor;
                if (isRTL) {// && alignNode == "r") {
                    marLStr = "padding-right:";// "margin-right: ";
                } else {
                    marLStr = "padding-left:";//"margin-left: ";
                }
                marRStr += ((marginRight + indent < 0) ? 0 : (marginRight + indent)) + "px;";
            }

            if (buType != "TYPE_NONE") {
                //var buFontAttrs = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs"]);
            }
            //console.log("Bullet Type: " + buType);
            //console.log("NumericTypr: " + buNum);
            //console.log("buChar: " + (buChar === undefined?'':buChar.charCodeAt(0)));
            //get definde bullet COLOR
            if (buClrNode === undefined){
                //lstStyle
                buClrNode = PPTXXmlUtils.getTextByPathList(lstStyle, [lvlStr, "a:buClr"]);
            }
            if (buClrNode === undefined) {
                buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buClr"]);
                if (buClrNode === undefined) {
                    buClrNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buClr"]);
                }
            }
            var defBultColor;
            if (buClrNode !== undefined) {
                defBultColor = PPTXStyleUtils.getSolidFill(buClrNode, undefined, undefined, warpObj);
            } else {
                if (pFontStyle !== undefined) {
                    //console.log("genBuChar pFontStyle: ", pFontStyle)
                    defBultColor = PPTXStyleUtils.getSolidFill(pFontStyle, undefined, undefined, warpObj);
                }
            }
            if (defBultColor === undefined || defBultColor == "NONE") {
                bultColor = dfltBultColor;
            } else {
                bultColor = [defBultColor, "", "solid"];
                color_tye = "solid";
            }
            //console.log("genBuChar node:", node, "pPrNode", pPrNode, " buClrNode: ", buClrNode, "defBultColor:", defBultColor,"dfltBultColor:" , dfltBultColor , "bultColor:", bultColor)

            //console.log("genBuChar: buClrNode: ", buClrNode, "bultColor", bultColor)
            //get definde bullet SIZE
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        var prcnt = parseInt(buFontSize) / 100000;
                        //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        var dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                }else{
                    bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
                }
            }
            if (buFontSize === undefined) {
                buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPts", "attrs", "val"]);
                if (buFontSize === undefined) {
                    buFontSize = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:buSzPct", "attrs", "val"]);
                    if (buFontSize !== undefined) {
                        var prcnt = parseInt(buFontSize) / 100000;
                        //dfltBultSize = XXpt
                        //var dfltBultSizeNoPt = dfltBultSize.substr(0, dfltBultSize.length - 2);
                        var dfltBultSizeNoPt = parseInt(dfltBultSize, "px");
                        bultSize = prcnt * (parseInt(dfltBultSizeNoPt)) + "px";// + "pt";
                    }
                } else {
                    bultSize = (parseInt(buFontSize) / 100) * fontSizeFactor + "px";
                }
            }
            if (buFontSize === undefined) {
                bultSize = dfltBultSize;
            }
            font_val = parseInt(bultSize, "px");
            ////////////////////////////////////////////////////////////////////////
            if (buType == "TYPE_BULLET") {
                var typefaceNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["a:buFont", "attrs", "typeface"]);
                var typeface = "";
                if (typefaceNode !== undefined) {
                    typeface = "font-family: " + typefaceNode;
                }
                // var marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(marLNode)) * slideFactor;
                // var marginRight = parseInt (PPTXXmlUtils.getTextByPathList(marRNode)) * slideFactor;
                // if (isNaN(marginLeft)) {
                //     marginLeft = 328600 * slideFactor;
                // }
                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }

                bullet = "<div style='height: 100%;" + typeface + ";" +
                    marLStr + marRStr +
                    "font-size:" + bultSize + ";" ;
                
                //bullet += "display: table-cell;";
                //"line-height: 0px;";
                if (color_tye == "solid") {
                    if (bultColor[0] !== undefined && bultColor[0] != "") {
                        bullet += "color:#" + bultColor[0] + "; ";
                    }
                    if (bultColor[1] !== undefined && bultColor[1] != "" && bultColor[1] != ";") {
                        bullet += "text-shadow:" + bultColor[1] + ";";
                    }
                    //no highlight/background-color to bullet
                    // if (bultColor[3] !== undefined && bultColor[3] != "") {
                    //     styleText += "background-color: #" + bultColor[3] + ";";
                    // }
                } else if (color_tye == "pattern" || color_tye == "pic" || color_tye == "gradient") {
                    if (color_tye == "pattern") {
                        bullet += "background:" + bultColor[0][0] + ";";
                        if (bultColor[0][1] !== null && bultColor[0][1] !== undefined && bultColor[0][1] != "") {
                            bullet += "background-size:" + bultColor[0][1] + ";";//" 2px 2px;" +
                        }
                        if (bultColor[0][2] !== null && bultColor[0][2] !== undefined && bultColor[0][2] != "") {
                            bullet += "background-position:" + bultColor[0][2] + ";";//" 2px 2px;" +
                        }
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";" +
                        //     "filter: " + bultColor[1].effcts + ";";
                    } else if (color_tye == "pic") {
                        bullet += bultColor[0] + ";";
                        // bullet += "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "color: transparent;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";

                    } else if (color_tye == "gradient") {

                        var colorAry = bultColor[0].color;
                        var rot = bultColor[0].rot;

                        bullet += "background: linear-gradient(" + rot + "deg,";
                        for (var i = 0; i < colorAry.length; i++) {
                            if (i == colorAry.length - 1) {
                                bullet += "#" + colorAry[i] + ");";
                            } else {
                                bullet += "#" + colorAry[i] + ", ";
                            }
                        }
                        // bullet += "color: transparent;" +
                        //     "-webkit-background-clip: text;" +
                        //     "background-clip: text;" +
                        //     "-webkit-text-stroke: " + bultColor[1].border + ";";
                    }
                    bullet += "-webkit-background-clip: text;" +
                        "background-clip: text;" +
                        "color: transparent;";
                    if (bultColor[1].border !== undefined && bultColor[1].border !== "") {
                        bullet += "-webkit-text-stroke: " + bultColor[1].border + ";";
                    }
                    if (bultColor[1].effcts !== undefined && bultColor[1].effcts !== "") {
                        bullet += "filter: " + bultColor[1].effcts + ";";
                    }
                }

                if (isRTL) {
                    //bullet += "display: inline-block;white-space: nowrap ;direction:rtl"; // float: right;  
                    bullet += "white-space: nowrap ;direction:rtl"; // display: table-cell;;
                }
                var isIE11 = !!window.MSInputMethodContext && !!document.documentMode;
                var htmlBu = buChar;

                if (!isIE11) {
                    //ie11 does not support unicode ?
                    htmlBu = getHtmlBullet(typefaceNode, buChar);
                }
                bullet += "'><div style='line-height: " + (font_val/2) + "px;'>" + htmlBu + "</div></div>"; //font_val
                //} 
                // else {
                //     marginLeft = 328600 * slideFactor * lvl;

                //     bullet = "<div style='" + marLStr + "'>" + buChar + "</div>";
                // }
            } else if (buType == "TYPE_NUMERIC") { ///////////Amir///////////////////////////////
                //if (buFontAttrs !== undefined) {
                // var marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
                // var marginRight = parseInt(buFontAttrs["pitchFamily"]);

                // if (isNaN(marginLeft)) {
                //     marginLeft = 328600 * slideFactor;
                // }
                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }
                //var typeface = buFontAttrs["typeface"];

                bullet = "<div style='height: 100%;" + marLStr + marRStr +
                    "color:#" + bultColor[0] + ";" +
                    "font-size:" + bultSize + ";";// +
                //"line-height: 0px;";
                if (isRTL) {
                    bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; // float: right;
                } else {
                    bullet += "display: inline-block;white-space: nowrap ;direction:ltr;"; //float: left;
                }
                bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></div>";
                // } else {
                //     marginLeft = 328600 * slideFactor * lvl;
                //     bullet = "<div style='margin-left: " + marginLeft + "px;";
                //     if (isRTL) {
                //         bullet += " float: right; direction:rtl;";
                //     } else {
                //         bullet += " float: left; direction:ltr;";
                //     }
                //     bullet += "' data-bulltname = '" + buNum + "' data-bulltlvl = '" + lvl + "' class='numeric-bullet-style'></div>";
                // }

            } else if (buType == "TYPE_BULPIC") { //PIC BULLET
                // var marginLeft = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marL"])) * slideFactor;
                // var marginRight = parseInt (PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "marR"])) * slideFactor;

                // if (isNaN(marginRight)) {
                //     marginRight = 0;
                // }
                // //console.log("marginRight: "+marginRight)
                // //buPic
                // if (isNaN(marginLeft)) {
                //     marginLeft = 328600 * slideFactor;
                // } else {
                //     marginLeft = 0;
                // }
                //var buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip","a:extLst","a:ext","asvg:svgBlip" , "attrs", "r:embed"]);
                var buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                var svgPicPath = "";
                var buImg;
                if (buPicId !== undefined) {
                    //svgPicPath = warpObj["slideResObj"][buPicId]["target"];
                    //buImg = warpObj["zip"].file(svgPicPath).asText();
                    //}else{
                    //buPicId = PPTXXmlUtils.getTextByPathList(buPic, ["a:blip", "attrs", "r:embed"]);
                    var imgPath = (warpObj["slideResObj"][buPicId] !== undefined) ? warpObj["slideResObj"][buPicId]["target"] : undefined;
                    //console.log("imgPath: ", imgPath);
                    if (imgPath === undefined) {
                        console.warn("Bullet image reference not found for buPicId:", buPicId);
                        buImg = "";
                    } else {
                        var imgFile = warpObj["zip"].file(imgPath);
                        if (imgFile === null) {
                            console.warn("Bullet image file not found:", imgPath);
                            buImg = "";
                        } else {
                            var imgArrayBuffer = imgFile.asArrayBuffer();
                            var imgExt = imgPath.split(".").pop();
                            var imgMimeType = PPTXXmlUtils.getMimeType(imgExt);
                            buImg = "<img src='data:" + imgMimeType + ";base64," + PPTXXmlUtils.base64ArrayBuffer(imgArrayBuffer) + "' style='width: 100%;'/>"// height: 100%
                            //console.log("imgPath: "+imgPath+"\nimgMimeType: "+imgMimeType)
                        }
                    }
                }
                if (buPicId === undefined) {
                    buImg = "&#8227;";
                }
                bullet = "<div style='height: 100%;" + marLStr + marRStr +
                    "width:" + bultSize + ";display: inline-block; ";// +
                //"line-height: 0px;";
                if (isRTL) {
                    bullet += "display: inline-block;white-space: nowrap ;direction:rtl;"; //direction:rtl; float: right;
                }
                bullet += "'>" + buImg + "  </div>";
                //////////////////////////////////////////////////////////////////////////////////////
            }
            // else {
            //     bullet = "<div style='margin-left: " + 328600 * slideFactor * lvl + "px" +
            //         "; margin-right: " + 0 + "px;'></div>";
            // }
            //console.log("genBuChar: width: ", $(bullet).outerWidth())
            return [bullet, margin_val, font_val];//$(bullet).outerWidth()];
        }
        function getHtmlBullet(typefaceNode, buChar) {
            //http://www.alanwood.net/demos/wingdings.html
            //not work for IE11
            //console.log("genBuChar typefaceNode:", typefaceNode, " buChar:", buChar, "charCodeAt:", buChar.charCodeAt(0))
            switch (buChar) {
                case "§":
                    return "&#9632;";//"■"; //9632 | U+25A0 | Black square
                    break;
                case "q":
                    return "&#10065;";//"❑"; // 10065 | U+2751 | Lower right shadowed white square
                    break;
                case "v":
                    return "&#10070;";//"❖"; //10070 | U+2756 | Black diamond minus white X
                    break;
                case "Ø":
                    return "&#11162;";//"⮚"; //11162 | U+2B9A | Three-D top-lighted rightwards equilateral arrowhead
                    break;
                case "ü":
                    return "&#10004;";//"✔";  //10004 | U+2714 | Heavy check mark
                    break;
                default:
                    if (/*typefaceNode == "Wingdings" ||*/ typefaceNode == "Wingdings 2" || typefaceNode == "Wingdings 3"){
                        var wingCharCode =  getDingbatToUnicode(typefaceNode, buChar);
                        if (wingCharCode !== null){
                            return "&#" + wingCharCode + ";";
                        }
                    }
                    return "&#" + (buChar.charCodeAt(0)) + ";";
            }
        }
        function getDingbatToUnicode(typefaceNode, buChar){
            if (dingbat_unicode){
                var dingbat_code = buChar.codePointAt(0) & 0xFFF;
                var char_unicode = null;
                var len = dingbat_unicode.length;
                var i = 0;
                while (len--) {
                    // blah blah
                    var item = dingbat_unicode[i];
                    if (item.f == typefaceNode && item.code == dingbat_code) {
                        char_unicode = item.unicode;
                        break;
                    }
                    i++;
                }
                return char_unicode
        }
    }

    /**
     * alphaNumeric - 将数字转换为字母数字格式
     * @param {number} num - 数字
     * @param {string} upperLower - 大小写选项（upperCase或lowerCase）
     * @returns {string} 字母数字格式的字符串
     */
    function alphaNumeric(num, upperLower) {
        num = Number(num) - 1;
        var aNum = "";
        if (upperLower == "upperCase") {
            aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toUpperCase();
        } else if (upperLower == "lowerCase") {
            aNum = (((num / 26 >= 1) ? String.fromCharCode(num / 26 + 64) : '') + String.fromCharCode(num % 26 + 65)).toLowerCase();
        }
        return aNum;
    }

    /**
     * archaicNumbers - 处理古数字格式
     * @param {Array} arr - 数字映射数组
     * @returns {Object} 包含format方法的对象
     */
    function archaicNumbers(arr) {
        var arrParse = arr.slice().sort(function (a, b) { return b[1].length - a[1].length });
        return {
            format: function (n) {
                var ret = '';
                for (var i = 0; i < arr.length; i++) {
                    var num = arr[i][0];
                    if (parseInt(num) > 0) {
                        for (; n >= num; n -= num) ret += arr[i][1];
                    } else {
                        ret = ret.replace(num, arr[i][1]);
                    }
                }
                return ret;
            }
        }
    }

    /**
     * romanize - 将数字转换为罗马数字
     * @param {number} num - 数字
     * @returns {string} 罗马数字字符串
     */
    function romanize(num) {
        if (!+num)
            return false;
        var digits = String(+num).split(""),
            key = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM",
                "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC",
                "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"],
            roman = "",
            i = 3;
        while (i--)
            roman = (key[+digits.pop() + (i * 10)] || "") + roman;
        return Array(+digits.join("") + 1).join("M") + roman;
    }
    var hebrew2Minus = archaicNumbers([
            [1000, ''],
            [400, 'ת'],
            [300, 'ש'],
            [200, 'ר'],
            [100, 'ק'],
            [90, 'צ'],
            [80, 'פ'],
            [70, 'ע'],
            [60, 'ס'],
            [50, 'נ'],
            [40, 'מ'],
            [30, 'ל'],
            [20, 'כ'],
            [10, 'י'],
            [9, 'ט'],
            [8, 'ח'],
            [7, 'ז'],
            [6, 'ו'],
            [5, 'ה'],
            [4, 'ד'],
            [3, 'ג'],
            [2, 'ב'],
            [1, 'א'],
            [/יה/, 'ט״ו'],
            [/יו/, 'ט״ז'],
            [/([א-ת])([א-ת])$/, '$1״$2'],
            [/^([א-ת])$/, "$1׳"]
        ]);
    /**
     * getNumTypeNum - 根据数字类型获取格式化的数字
     * @param {string} numTyp - 数字类型
     * @param {number} num - 数字
     * @returns {string} 格式化的数字字符串
     */
    function getNumTypeNum(numTyp, num) {
        var rtrnNum = "";
        switch (numTyp) {
            case "arabicPeriod":
                rtrnNum = num + ". ";
                break;
            case "arabicParenR":
                rtrnNum = num + ") ";
                break;
            case "alphaLcParenR":
                rtrnNum = alphaNumeric(num, "lowerCase") + ") ";
                break;
            case "alphaLcPeriod":
                rtrnNum = alphaNumeric(num, "lowerCase") + ". ";
                break;

            case "alphaUcParenR":
                rtrnNum = alphaNumeric(num, "upperCase") + ") ";
                break;
            case "alphaUcPeriod":
                rtrnNum = alphaNumeric(num, "upperCase") + ". ";
                break;

            case "romanUcPeriod":
                rtrnNum = romanize(num) + ". ";
                break;
            case "romanLcParenR":
                rtrnNum = romanize(num) + ") ";
                break;
            case "hebrew2Minus":
                rtrnNum = hebrew2Minus.format(num) + "-";
                break;
            default:
                rtrnNum = num;
        }
        return rtrnNum;
    }

    function genSpanElement(node, rIndex, pNode, textBodyNode, pFontStyle, slideLayoutSpNode, idx, type, rNodeLength, warpObj, isBullate) {
            //https://codepen.io/imdunn/pen/GRgwaye ?
            var text_style = "";
            var lstStyle = textBodyNode["a:lstStyle"];
            var slideMasterTextStyles = warpObj["slideMasterTextStyles"];

            var text = node["a:t"];
            //var text_count = text.length;

            var openElemnt = "<span";//"<bdi";
            var closeElemnt = "</span>";// "</bdi>";
            var styleText = "";
            if (text === undefined && node["type"] !== undefined) {
                if (is_first_br) {
                    //openElemnt = "<br";
                    //closeElemnt = "";
                    //return "<br style='font-size: initial'>"
                    is_first_br = false;
                    return "<span class='line-break-br' ></span>";
                } else {
                    // styleText += "display: block;";
                    // openElemnt = "<span";
                    // closeElemnt = "</span>";
                }

                styleText += "display: block;";
                //openElemnt = "<span";
                //closeElemnt = "</span>";
            } else {

                is_first_br = true;
            }
            if (typeof text !== 'string') {
                text = PPTXXmlUtils.getTextByPathList(node, ["a:fld", "a:t"]);
                if (typeof text !== 'string') {
                    text = "&nbsp;";
                    //return "<span class='text-block '>&nbsp;</span>";
                }
                // if (text === undefined) {
                //     return "";
                // }
            }

            var pPrNode = pNode["a:pPr"];
            //lvl
            var lvl = 1;
            var lvlNode = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "lvl"]);
            if (lvlNode !== undefined) {
                lvl = parseInt(lvlNode) + 1;
            }
            //console.log("genSpanElement node: ", node, "rIndex: ", rIndex, ", pNode: ", pNode, ",pPrNode: ", pPrNode, "pFontStyle:", pFontStyle, ", idx: ", idx, "type:", type, warpObj);
            var layoutMasterNode = PPTXStyleUtils.getLayoutAndMasterNode(pNode, idx, type, warpObj);
            var pPrNodeLaout = layoutMasterNode.nodeLaout;
            var pPrNodeMaster = layoutMasterNode.nodeMaster;

            //Language
            var lang = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "lang"]);
            var isRtlLan = (lang !== undefined && rtl_langs_array.indexOf(lang) !== -1)?true:false;
            //rtl
            var getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNode, ["attrs", "rtl"]);
            if (getRtlVal === undefined) {
                getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["attrs", "rtl"]);
                if (getRtlVal === undefined && type != "shape") {
                    getRtlVal = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["attrs", "rtl"]);
                }
            }
            var isRTL = false;
            var dirStr = "ltr";
            if (getRtlVal !== undefined && getRtlVal == "1") {
                isRTL = true;
                dirStr = "rtl";
            }

            var linkID = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "r:id"]);
            var linkTooltip = "";
            var defLinkClr;
            if (linkID !== undefined) {
                linkTooltip = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:hlinkClick", "attrs", "tooltip"]);
                if (linkTooltip !== undefined) {
                    linkTooltip = "title='" + linkTooltip + "'";
                }
                defLinkClr = PPTXStyleUtils.getSchemeColorFromTheme("a:hlink", undefined, undefined, warpObj);

                var linkClrNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);// PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:solidFill"]);
                var rPrlinkClr = PPTXStyleUtils.getSolidFill(linkClrNode, undefined, undefined, warpObj);


                //console.log("genSpanElement defLinkClr: ", defLinkClr, "rPrlinkClr:", rPrlinkClr)
                if (rPrlinkClr !== undefined && rPrlinkClr != "") {
                    defLinkClr = rPrlinkClr;
                }

            }
            /////////////////////////////////////////////////////////////////////////////////////
            //getFontColor
            var fontClrPr = PPTXStyleUtils.getFontColorPr(node, pNode, lstStyle, pFontStyle, lvl, idx, type, warpObj);
            var fontClrType = fontClrPr[2];
            //console.log("genSpanElement fontClrPr: ", fontClrPr, "linkID", linkID);
            if (fontClrType == "solid") {
                if (linkID === undefined && fontClrPr[0] !== undefined && fontClrPr[0] != "") {
                    styleText += "color: #" + fontClrPr[0] + ";";
                }
                else if (linkID !== undefined && defLinkClr !== undefined) {
                    styleText += "color: #" + defLinkClr + ";";
                }

                if (fontClrPr[1] !== undefined && fontClrPr[1] != "" && fontClrPr[1] != ";") {
                    styleText += "text-shadow:" + fontClrPr[1] + ";";
                }
                if (fontClrPr[3] !== undefined && fontClrPr[3] != "") {
                    styleText += "background-color: #" + fontClrPr[3] + ";";
                }
            } else if (fontClrType == "pattern" || fontClrType == "pic" || fontClrType == "gradient") {
                if (fontClrType == "pattern") {
                    styleText += "background:" + fontClrPr[0][0] + ";";
                    if (fontClrPr[0][1] !== null && fontClrPr[0][1] !== undefined && fontClrPr[0][1] != "") {
                        styleText += "background-size:" + fontClrPr[0][1] + ";";//" 2px 2px;" +
                    }
                    if (fontClrPr[0][2] !== null && fontClrPr[0][2] !== undefined && fontClrPr[0][2] != "") {
                        styleText += "background-position:" + fontClrPr[0][2] + ";";//" 2px 2px;" +
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";" +
                    //     "filter: " + fontClrPr[1].effcts + ";";
                } else if (fontClrType == "pic") {
                    styleText += fontClrPr[0] + ";";
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";
                } else if (fontClrType == "gradient") {

                    var colorAry = fontClrPr[0].color;
                    var rot = fontClrPr[0].rot;

                    styleText += "background: linear-gradient(" + rot + "deg,";
                    for (var i = 0; i < colorAry.length; i++) {
                        if (i == colorAry.length - 1) {
                            styleText += "#" + colorAry[i] + ");";
                        } else {
                            styleText += "#" + colorAry[i] + ", ";
                        }
                    }
                    // styleText += "-webkit-background-clip: text;" +
                    //     "background-clip: text;" +
                    //     "color: transparent;" +
                    //     "-webkit-text-stroke: " + fontClrPr[1].border + ";";

                }
                styleText += "-webkit-background-clip: text;" +
                    "background-clip: text;" +
                    "color: transparent;";
                if (fontClrPr[1].border !== undefined && fontClrPr[1].border !== "") {
                    styleText += "-webkit-text-stroke: " + fontClrPr[1].border + ";";
                }
                if (fontClrPr[1].effcts !== undefined && fontClrPr[1].effcts !== "") {
                    styleText += "filter: " + fontClrPr[1].effcts + ";";
                }
            }
            var font_size = PPTXStyleUtils.getFontSize(node, textBodyNode, pFontStyle, lvl, type, warpObj);
            //text_style += "font-size:" + font_size + ";"
            
            text_style += "font-size:" + font_size + ";" +
                // marLStr +
                "font-family:" + PPTXStyleUtils.getFontType(node, type, warpObj, pFontStyle) + ";" +
                "font-weight:" + PPTXStyleUtils.getFontBold(node, type, slideMasterTextStyles) + ";" +
                "font-style:" + PPTXStyleUtils.getFontItalic(node, type, slideMasterTextStyles) + ";" +
                "text-decoration:" + PPTXStyleUtils.getFontDecoration(node, type, slideMasterTextStyles) + ";" +
                "text-align:" + PPTXStyleUtils.getTextHorizontalAlign(node, pNode, type, warpObj) + ";" +
                "vertical-align:" + PPTXStyleUtils.getTextVerticalAlign(node, type, slideMasterTextStyles) + ";";
            //rNodeLength
            //console.log("genSpanElement node:", node, "lang:", lang, "isRtlLan:", isRtlLan, "span parent dir:", dirStr)
            if (isRtlLan) { //|| rIndex === undefined
                styleText += "direction:rtl;";
            }else{ //|| rIndex === undefined
                styleText += "direction:ltr;";
            }
            // } else if (dirStr == "rtl" && isRtlLan ) {
            //     styleText += "direction:rtl;";

            // } else if (dirStr == "ltr" && !isRtlLan ) {
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "ltr" && isRtlLan){
            //     styleText += "direction:ltr;";
            // }else{
            //     styleText += "direction:inherit;";
            // }

            // if (dirStr == "rtl" && !isRtlLan) { //|| rIndex === undefined
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "rtl" && isRtlLan) {
            //     styleText += "direction:rtl;";
            // } else if (dirStr == "ltr" && !isRtlLan) {
            //     styleText += "direction:ltr;";
            // } else if (dirStr == "ltr" && isRtlLan) {
            //     styleText += "direction:rtl;";
            // } else {
            //     styleText += "direction:inherit;";
            // }

            //     //"direction:" + dirStr + ";";
            //if (rNodeLength == 1 || rIndex == 0 ){
            //styleText += "display: table-cell;white-space: nowrap;";
            //}
            var highlight = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "a:highlight"]);
            if (highlight !== undefined) {
                styleText += "background-color:#" + PPTXStyleUtils.getSolidFill(highlight, undefined, undefined, warpObj) + ";";
                //styleText += "Opacity:" + getColorOpacity(highlight) + ";";
            }

            //letter-spacing:
            var spcNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "spc"]);
            if (spcNode === undefined) {
                spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "spc"]);
                if (spcNode === undefined) {
                    spcNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "spc"]);
                }
            }
            if (spcNode !== undefined) {
                var ltrSpc = parseInt(spcNode) / 100; //pt
                styleText += "letter-spacing: " + ltrSpc + "px;";// + "pt;";
            }

            //Text Cap Types
            var capNode = PPTXXmlUtils.getTextByPathList(node, ["a:rPr", "attrs", "cap"]);
            if (capNode === undefined) {
                capNode = PPTXXmlUtils.getTextByPathList(pPrNodeLaout, ["a:defRPr", "attrs", "cap"]);
                if (capNode === undefined) {
                    capNode = PPTXXmlUtils.getTextByPathList(pPrNodeMaster, ["a:defRPr", "attrs", "cap"]);
                }
            }
            if (capNode == "small" || capNode == "all") {
                styleText += "text-transform: uppercase";
            }
            //styleText += "word-break: break-word;";
            //console.log("genSpanElement node: ", node, ", capNode: ", capNode, ",pPrNodeLaout: ", pPrNodeLaout, ", pPrNodeMaster: ", pPrNodeMaster, "warpObj:", warpObj);

            var cssName = "";

            if (styleText in warpObj.styleTable) {
                cssName = warpObj.styleTable[styleText]["name"];
            } else {
                cssName = "_css_" + (Object.keys(warpObj.styleTable).length + 1);
                warpObj.styleTable[styleText] = {
                    "name": cssName,
                    "text": styleText
                };
            }
            var linkColorSyle = "";
            if (fontClrType == "solid" && linkID !== undefined) {
                linkColorSyle = "style='color: inherit;'";
            }

            if (linkID !== undefined && linkID != "") {
                var linkURL = warpObj["slideResObj"][linkID]["target"];
                linkURL = PPTXXmlUtils.escapeHtml(linkURL);
                return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'><a href='" + linkURL + "' " + linkColorSyle + "  " + linkTooltip + " target='_blank'>" +
                        text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + "</a>" + closeElemnt;
            } else {
                return openElemnt + " class='text-block " + cssName + "' style='" + text_style + "'>" + text.replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;').replace(/\s/g, "&nbsp;") + closeElemnt;//"</bdi>";
            }

        }

    
        function genChart(node, warpObj) {

            var order = node["attrs"]["order"];
            var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
            var result = "<div id='chart" + warpObj.chartID + "' class='block content' style='" +
                PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined) + PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
                " z-index: " + order + ";'></div>";

            var rid = node["a:graphic"]["a:graphicData"]["c:chart"]["attrs"]["r:id"];
            var refName = warpObj["slideResObj"][rid]["target"];
            var content = PPTXXmlUtils.readXmlFile(warpObj["zip"], refName);
            var plotArea = PPTXXmlUtils.getTextByPathList(content, ["c:chartSpace", "c:chart", "c:plotArea"]);

            var chartData = null;
            for (var key in plotArea) {
                switch (key) {
                    case "c:lineChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "lineChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:barChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "barChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:pieChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "pieChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:pie3DChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "pie3DChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:areaChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "areaChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:scatterChart":
                        chartData = {
                            "type": "createChart",
                            "data": {
                                "chartID": "chart" + warpObj.chartID,
                                "chartType": "scatterChart",
                                "chartData": PPTXStyleUtils.extractChartData(plotArea[key]["c:ser"])
                            }
                        };
                        break;
                    case "c:catAx":
                        break;
                    case "c:valAx":
                        break;
                    default:
                }
            }

            if (chartData !== null) {
                warpObj.MsgQueue.push(chartData);
            }

            warpObj.chartID++;
            return result;
        }

        function genTable(node, warpObj) {
            var order = node["attrs"]["order"];
            var tableNode = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl"]);
            var xfrmNode = PPTXXmlUtils.getTextByPathList(node, ["p:xfrm"]);
            /////////////////////////////////////////Amir////////////////////////////////////////////////
            var getTblPr = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblPr"]);
            var getColsGrid = PPTXXmlUtils.getTextByPathList(node, ["a:graphic", "a:graphicData", "a:tbl", "a:tblGrid", "a:gridCol"]);
            var tblDir = "";
            if (getTblPr !== undefined) {
                var isRTL = getTblPr["attrs"]["rtl"];
                tblDir = (isRTL == 1 ? "dir=rtl" : "dir=ltr");
            }
            var firstRowAttr = getTblPr["attrs"]["firstRow"]; //associated element <a:firstRow> in the table styles
            var firstColAttr = getTblPr["attrs"]["firstCol"]; //associated element <a:firstCol> in the table styles
            var lastRowAttr = getTblPr["attrs"]["lastRow"]; //associated element <a:lastRow> in the table styles
            var lastColAttr = getTblPr["attrs"]["lastCol"]; //associated element <a:lastCol> in the table styles
            var bandRowAttr = getTblPr["attrs"]["bandRow"]; //associated element <a:band1H>, <a:band2H> in the table styles
            var bandColAttr = getTblPr["attrs"]["bandCol"]; //associated element <a:band1V>, <a:band2V> in the table styles
            //console.log("getTblPr: ", getTblPr);
            var tblStylAttrObj = {
                isFrstRowAttr: (firstRowAttr !== undefined && firstRowAttr == "1") ? 1 : 0,
                isFrstColAttr: (firstColAttr !== undefined && firstColAttr == "1") ? 1 : 0,
                isLstRowAttr: (lastRowAttr !== undefined && lastRowAttr == "1") ? 1 : 0,
                isLstColAttr: (lastColAttr !== undefined && lastColAttr == "1") ? 1 : 0,
                isBandRowAttr: (bandRowAttr !== undefined && bandRowAttr == "1") ? 1 : 0,
                isBandColAttr: (bandColAttr !== undefined && bandColAttr == "1") ? 1 : 0
            }

            var thisTblStyle;
            var tbleStyleId = getTblPr["a:tableStyleId"];
            if (tbleStyleId !== undefined) {
                var tbleStylList = warpObj.tableStyles["a:tblStyleLst"]["a:tblStyle"];
                if (tbleStylList !== undefined) {
                    if (tbleStylList.constructor === Array) {
                        for (var k = 0; k < tbleStylList.length; k++) {
                            if (tbleStylList[k]["attrs"]["styleId"] == tbleStyleId) {
                                thisTblStyle = tbleStylList[k];
                            }
                        }
                    } else {
                        if (tbleStylList["attrs"]["styleId"] == tbleStyleId) {
                            thisTblStyle = tbleStylList;
                        }
                    }
                }
            }
            if (thisTblStyle !== undefined) {
                thisTblStyle["tblStylAttrObj"] = tblStylAttrObj;
                warpObj["thisTbiStyle"] = thisTblStyle;
            }
            var tblStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle"]);
            var tblBorderStyl = PPTXXmlUtils.getTextByPathList(tblStyl, ["a:tcBdr"]);
            var tbl_borders = "";
            if (tblBorderStyl !== undefined) {
                tbl_borders = PPTXStyleUtils.getTableBorders(tblBorderStyl, warpObj);
            }
            var tbl_bgcolor = "";
            var tbl_opacity = 1;
            var tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:tblBg", "a:fillRef"]);
            //console.log( "thisTblStyle:", thisTblStyle, "warpObj:", warpObj)
            if (tbl_bgFillschemeClr !== undefined) {
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            if (tbl_bgFillschemeClr === undefined) {
                tbl_bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                tbl_bgcolor = PPTXStyleUtils.getSolidFill(tbl_bgFillschemeClr, undefined, undefined, warpObj);
            }
            if (tbl_bgcolor !== "") {
                tbl_bgcolor = "background-color: #" + tbl_bgcolor + ";";
            }
            ////////////////////////////////////////////////////////////////////////////////////////////
            var tableHtml = "<table " + tblDir + " style='border-collapse: collapse;" +
                PPTXXmlUtils.getPosition(xfrmNode, node, undefined, undefined) +
                PPTXXmlUtils.getSize(xfrmNode, undefined, undefined) +
                " z-index: " + order + ";" +
                tbl_borders + ";" +
                tbl_bgcolor + "'>";

            var trNodes = tableNode["a:tr"];
            if (trNodes.constructor !== Array) {
                trNodes = [trNodes];
            }
            //if (trNodes.constructor === Array) {
                //multi rows
                var totalrowSpan = 0;
                var rowSpanAry = [];
                for (var i = 0; i < trNodes.length; i++) {
                    //////////////rows Style ////////////Amir
                    var rowHeightParam = trNodes[i]["attrs"]["h"];
                    var rowHeight = 0;
                    var rowsStyl = "";
                    if (rowHeightParam !== undefined) {
                        rowHeight = parseInt(rowHeightParam) * slideFactor;
                        rowsStyl += "height:" + rowHeight + "px;";
                    }
                    var fillColor = "";
                    var row_borders = "";
                    var fontClrPr = "";
                    var fontWeight = "";
                    var band_1H_fillColor;
                    var band_2H_fillColor;

                    if (thisTblStyle !== undefined && thisTblStyle["a:wholeTbl"] !== undefined) {
                        var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontColor = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontColor !== undefined) {
                                fontClrPr = local_fontColor;
                            }

                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight != "") {
                                fontWeight = local_fontWeight
                            }
                        }
                    }

                    if (i == 0 && tblStylAttrObj["isFrstRowAttr"] == 1 && thisTblStyle !== undefined) {

                        var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                        }
                        var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:firstRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }
                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }

                    } else if (i > 0 && tblStylAttrObj["isBandRowAttr"] == 1 && thisTblStyle !== undefined) {
                        fillColor = "";
                        row_borders = undefined;
                        if ((i % 2) == 0 && thisTblStyle["a:band2H"] !== undefined) {
                            // console.log("i: ", i, 'thisTblStyle["a:band2H"]:', thisTblStyle["a:band2H"])
                            //check if there is a row bg
                            var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== "") {
                                    fillColor = local_fillColor;
                                    band_2H_fillColor = local_fillColor;
                                }
                            }


                            var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                            }

                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");

                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                        if ((i % 2) != 0 && thisTblStyle["a:band1H"] !== undefined) {
                            var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:fill", "a:solidFill"]);
                            if (bgFillschemeClr !== undefined) {
                                var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                                if (local_fillColor !== undefined) {
                                    fillColor = local_fillColor;
                                    band_1H_fillColor = local_fillColor;
                                }
                            }
                            var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcStyle", "a:tcBdr"]);
                            if (borderStyl !== undefined) {
                                var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                                if (local_row_borders != "") {
                                    row_borders = local_row_borders;
                                }
                            }
                            var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1H", "a:tcTxStyle"]);
                            if (rowTxtStyl !== undefined) {
                                var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                                if (local_fontClrPr !== undefined) {
                                    fontClrPr = local_fontClrPr;
                                }
                                var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                                if (local_fontWeight != "") {
                                    fontWeight = local_fontWeight;
                                }
                            }
                        }

                    }
                    //last row
                    if (i == (trNodes.length - 1) && tblStylAttrObj["isLstRowAttr"] == 1 && thisTblStyle !== undefined) {
                        var bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:fill", "a:solidFill"]);
                        if (bgFillschemeClr !== undefined) {
                            var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                            if (local_fillColor !== undefined) {
                                fillColor = local_fillColor;
                            }
                            // var local_colorOpacity = getColorOpacity(bgFillschemeClr);
                            // if(local_colorOpacity !== undefined){
                            //     colorOpacity = local_colorOpacity;
                            // }
                        }
                        var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcStyle", "a:tcBdr"]);
                        if (borderStyl !== undefined) {
                            var local_row_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
                            if (local_row_borders != "") {
                                row_borders = local_row_borders;
                            }
                        }
                        var rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:lastRow", "a:tcTxStyle"]);
                        if (rowTxtStyl !== undefined) {
                            var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                            if (local_fontClrPr !== undefined) {
                                fontClrPr = local_fontClrPr;
                            }

                            var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                            if (local_fontWeight !== "") {
                                fontWeight = local_fontWeight;
                            }
                        }
                    }
                    rowsStyl += ((row_borders !== undefined) ? row_borders : "");
                    rowsStyl += ((fontClrPr !== undefined) ? " color: #" + fontClrPr + ";" : "");
                    rowsStyl += ((fontWeight != "") ? " font-weight:" + fontWeight + ";" : "");
                    if (fillColor !== undefined && fillColor != "") {
                        //rowsStyl += "background-color: rgba(" + hexToRgbNew(fillColor) + "," + colorOpacity + ");";
                        rowsStyl += "background-color: #" + fillColor + ";";
                    }
                    tableHtml += "<tr style='" + rowsStyl + "'>";
                    ////////////////////////////////////////////////

                    var tcNodes = trNodes[i]["a:tc"];
                    if (tcNodes !== undefined) {
                        if (tcNodes.constructor === Array) {
                            //multi columns
                            var j = 0;
                            if (rowSpanAry.length == 0) {
                                rowSpanAry = Array.apply(null, Array(tcNodes.length)).map(function () { return 0 });
                            }
                            var totalColSpan = 0;
                            while (j < tcNodes.length) {
                                if (rowSpanAry[j] == 0 && totalColSpan == 0) {
                                    var a_sorce;
                                    //j=0 : first col
                                    if (j == 0 && tblStylAttrObj["isFrstColAttr"] == 1) {
                                        a_sorce = "a:firstCol";
                                        if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) &&
                                            PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:seCell"]) !== undefined) {
                                            a_sorce = "a:seCell";
                                        } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 &&
                                            PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:neCell"]) !== undefined) {
                                            a_sorce = "a:neCell";
                                        }
                                    } else if ((j > 0 && tblStylAttrObj["isBandColAttr"] == 1) &&
                                        !(tblStylAttrObj["isFrstColAttr"] == 1 && i == 0) &&
                                        !(tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1)) &&
                                        j != (tcNodes.length - 1)) {

                                        if ((j % 2) != 0) {

                                            var aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                            if (aBandNode === undefined) {
                                                aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                                if (aBandNode !== undefined) {
                                                    a_sorce = "a:band2V";
                                                }
                                            } else {
                                                a_sorce = "a:band2V";
                                            }

                                        }
                                    }

                                    if (j == (tcNodes.length - 1) && tblStylAttrObj["isLstColAttr"] == 1) {
                                        a_sorce = "a:lastCol";
                                        if (tblStylAttrObj["isLstRowAttr"] == 1 && i == (trNodes.length - 1) && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:swCell"]) !== undefined) {
                                            a_sorce = "a:swCell";
                                        } else if (tblStylAttrObj["isFrstRowAttr"] == 1 && i == 0 && PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:nwCell"]) !== undefined) {
                                            a_sorce = "a:nwCell";
                                        }
                                    }

                                    var cellParmAry = getTableCellParams(tcNodes[j], getColsGrid, i , j , thisTblStyle, a_sorce, warpObj)
                                    var text = cellParmAry[0];
                                    var colStyl = cellParmAry[1];
                                    var cssName = cellParmAry[2];
                                    var rowSpan = cellParmAry[3];
                                    var colSpan = cellParmAry[4];



                                    if (rowSpan !== undefined) {
                                        totalrowSpan++;
                                        rowSpanAry[j] = parseInt(rowSpan) - 1;
                                        tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' rowspan ='" +
                                            parseInt(rowSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                    } else if (colSpan !== undefined) {
                                        tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' colspan = '" +
                                            parseInt(colSpan) + "' style='" + colStyl + "'>" + text + "</td>";
                                        totalColSpan = parseInt(colSpan) - 1;
                                    } else {
                                        tableHtml += "<td class='" + cssName + "' data-row='" + i + "," + j + "' style = '" + colStyl + "'>" + text + "</td>";
                                    }

                                } else {
                                    if (rowSpanAry[j] != 0) {
                                        rowSpanAry[j] -= 1;
                                    }
                                    if (totalColSpan != 0) {
                                        totalColSpan--;
                                    }
                                }
                                j++;
                            }
                        } else {
                            //single column 

                            var a_sorce;
                            if (tblStylAttrObj["isFrstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                                a_sorce = "a:firstCol";

                            } else if ((tblStylAttrObj["isBandColAttr"] == 1) && !(tblStylAttrObj["isLstRowAttr"] == 1)) {

                                var aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band2V"]);
                                if (aBandNode === undefined) {
                                    aBandNode = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:band1V"]);
                                    if (aBandNode !== undefined) {
                                        a_sorce = "a:band2V";
                                    }
                                } else {
                                    a_sorce = "a:band2V";
                                }
                            }

                            if (tblStylAttrObj["isLstColAttr"] == 1 && !(tblStylAttrObj["isLstRowAttr"] == 1)) {
                                a_sorce = "a:lastCol";
                            }


                            var cellParmAry = getTableCellParams(tcNodes, getColsGrid , i , undefined , thisTblStyle, a_sorce, warpObj)
                            var text = cellParmAry[0];
                            var colStyl = cellParmAry[1];
                            var cssName = cellParmAry[2];
                            var rowSpan = cellParmAry[3];

                            if (rowSpan !== undefined) {
                                tableHtml += "<td  class='" + cssName + "' rowspan='" + parseInt(rowSpan) + "' style = '" + colStyl + "'>" + text + "</td>";
                            } else {
                                tableHtml += "<td class='" + cssName + "' style='" + colStyl + "'>" + text + "</td>";
                            }
                        }
                    }
                    tableHtml += "</tr>";
                }
                //////////////////////////////////////////////////////////////////////////////////
            

            return tableHtml;
        }
        
        function getTableCellParams(tcNodes, getColsGrid , row_idx , col_idx , thisTblStyle, cellSource, warpObj) {
            //thisTblStyle["a:band1V"] => thisTblStyle[cellSource]
            //text, cell-width, cell-borders, 
            //var text = PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj);//tableStyles
            var rowSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "rowSpan"]);
            var colSpan = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "gridSpan"]);
            var vMerge = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "vMerge"]);
            var hMerge = PPTXXmlUtils.getTextByPathList(tcNodes, ["attrs", "hMerge"]);
            var colStyl = "word-wrap: break-word;";
            var colWidth;
            var celFillColor = "";
            var col_borders = "";
            var colFontClrPr = "";
            var colFontWeight = "";
            var lin_bottm = "",
                lin_top = "",
                lin_left = "",
                lin_right = "",
                lin_bottom_left_to_top_right = "",
                lin_top_left_to_bottom_right = "";
            
            var colSapnInt = parseInt(colSpan);
            var total_col_width = 0;
            if (!isNaN(colSapnInt) && colSapnInt > 1){
                for (var k = 0; k < colSapnInt ; k++) {
                    total_col_width += parseInt (PPTXXmlUtils.getTextByPathList(getColsGrid[col_idx + k], ["attrs", "w"]));
                }
            }else{
                total_col_width = PPTXXmlUtils.getTextByPathList((col_idx === undefined) ? getColsGrid : getColsGrid[col_idx], ["attrs", "w"]);
            }
            

            var text = PPTXTextUtils.genTextBody(tcNodes["a:txBody"], tcNodes, undefined, undefined, undefined, undefined, warpObj, total_col_width);//tableStyles

            if (total_col_width != 0 /*&& row_idx == 0*/) {
                colWidth = parseInt(total_col_width) * slideFactor;
                colStyl += "width:" + colWidth + "px;";
            }

            //cell bords
            lin_bottm = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnB"]);
            if (lin_bottm === undefined && cellSource !== undefined) {
                if (cellSource !== undefined)
                    lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
                if (lin_bottm === undefined) {
                    lin_bottm = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:bottom", "a:ln"]);
                }
            }
            lin_top = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnT"]);
            if (lin_top === undefined) {
                if (cellSource !== undefined)
                    lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
                if (lin_top === undefined) {
                    lin_top = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:top", "a:ln"]);
                }
            }
            lin_left = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnL"]);
            if (lin_left === undefined) {
                if (cellSource !== undefined)
                    lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
                if (lin_left === undefined) {
                    lin_left = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:left", "a:ln"]);
                }
            }
            lin_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnR"]);
            if (lin_right === undefined) {
                if (cellSource !== undefined)
                    lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle[cellSource], ["a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
                if (lin_right === undefined) {
                    lin_right = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcStyle", "a:tcBdr", "a:right", "a:ln"]);
                }
            }
            lin_bottom_left_to_top_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:lnBlToTr"]);
            lin_top_left_to_bottom_right = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr", "a:InTlToBr"]);

            if (lin_bottm !== undefined && lin_bottm != "") {
                var bottom_line_border = PPTXStyleUtils.getBorder(lin_bottm, undefined, false, "", warpObj)
                if (bottom_line_border != "") {
                    colStyl += "border-bottom:" + bottom_line_border + ";";
                }
            }
            if (lin_top !== undefined && lin_top != "") {
                var top_line_border = PPTXStyleUtils.getBorder(lin_top, undefined, false, "", warpObj);
                if (top_line_border != "") {
                    colStyl += "border-top: " + top_line_border + ";";
                }
            }
            if (lin_left !== undefined && lin_left != "") {
                var left_line_border = PPTXStyleUtils.getBorder(lin_left, undefined, false, "", warpObj)
                if (left_line_border != "") {
                    colStyl += "border-left: " + left_line_border + ";";
                }
            }
            if (lin_right !== undefined && lin_right != "") {
                var right_line_border = PPTXStyleUtils.getBorder(lin_right, undefined, false, "", warpObj)
                if (right_line_border != "") {
                    colStyl += "border-right:" + right_line_border + ";";
                }
            }

            //cell fill color custom
            var getCelFill = PPTXXmlUtils.getTextByPathList(tcNodes, ["a:tcPr"]);
            if (getCelFill !== undefined && getCelFill != "") {
                var cellObj = {
                    "p:spPr": getCelFill
                };
                celFillColor = PPTXStyleUtils.getShapeFill(cellObj, undefined, false, warpObj, "slide")
            }

            //cell fill color theme
            if (celFillColor == "" || celFillColor == "background-color: inherit;") {
                var bgFillschemeClr;
                if (cellSource !== undefined)
                    bgFillschemeClr = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:fill", "a:solidFill"]);
                if (bgFillschemeClr !== undefined) {
                    var local_fillColor = PPTXStyleUtils.getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
                    if (local_fillColor !== undefined) {
                        celFillColor = " background-color: #" + local_fillColor + ";";
                    }
                }
            }
            var cssName = "";
            if (celFillColor !== undefined && celFillColor != "") {
                if (celFillColor in warpObj.styleTable) {
                    cssName = warpObj.styleTable[celFillColor]["name"];
                } else {
                    cssName = "_tbl_cell_css_" + (Object.keys(warpObj.styleTable).length + 1);
                    warpObj.styleTable[celFillColor] = {
                        "name": cssName,
                        "text": celFillColor
                    };
                }

            }

            //border
            // var borderStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcStyle", "a:tcBdr"]);
            // if (borderStyl !== undefined) {
            //     var local_col_borders = PPTXStyleUtils.getTableBorders(borderStyl, warpObj);
            //     if (local_col_borders != "") {
            //         col_borders = local_col_borders;
            //     }
            // }
            // if (col_borders != "") {
            //     colStyl += col_borders;
            // }

            //Text style
            var rowTxtStyl;
            if (cellSource !== undefined) {
                rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, [cellSource, "a:tcTxStyle"]);
            }
            // if (rowTxtStyl === undefined) {
            //     rowTxtStyl = PPTXXmlUtils.getTextByPathList(thisTblStyle, ["a:wholeTbl", "a:tcTxStyle"]);
            // }
            if (rowTxtStyl !== undefined) {
                var local_fontClrPr = PPTXStyleUtils.getSolidFill(rowTxtStyl, undefined, undefined, warpObj);
                if (local_fontClrPr !== undefined) {
                    colFontClrPr = local_fontClrPr;
                }
                var local_fontWeight = ( (PPTXXmlUtils.getTextByPathList(rowTxtStyl, ["attrs", "b"]) == "on") ? "bold" : "");
                if (local_fontWeight !== "") {
                    colFontWeight = local_fontWeight;
                }
            }
            colStyl += ((colFontClrPr !== "") ? "color: #" + colFontClrPr + ";" : "");
            colStyl += ((colFontWeight != "") ? " font-weight:" + colFontWeight + ";" : "");

            return [text, colStyl, cssName, rowSpan, colSpan];
        }
    return {
        genTextBody,
        genBuChar,
        getHtmlBullet,
        getDingbatToUnicode,
        genSpanElement,
        genChart,
        genTable,
        getTableCellParams,
        alphaNumeric,
        archaicNumbers,
        romanize,
        getNumTypeNum,
    };
})();
