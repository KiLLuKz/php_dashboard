// Copyright 2006 Google Inc.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//   http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


// Known Issues:
//
// * Patterns only support repeat.
// * Radial gradient are not implemented. The VML version of these look very
//   different from the canvas one.
// * Clipping paths are not implemented.
// * Coordsize. The width and height attribute have higher priority than the
//   width and height style values which isn't correct.
// * Painting mode isn't implemented.
// * Canvas width/height should is using content-box by default. IE in
//   Quirks mode will draw the canvas using border-box. Either change your
//   doctype to HTML5
//   (http://www.whatwg.org/specs/web-apps/current-work/#the-doctype)
//   or use Box Sizing Behavior from WebFX
//   (http://webfx.eae.net/dhtml/boxsizing/boxsizing.html)
// * Non uniform scaling does not correctly scale strokes.
// * Filling very large shapes (above 5000 points) is buggy.
// * Optimize. There is always room for speed improvements.

// Only add this code if we do not already have a canvas implementation
if (!document.createElement('canvas').getContext) {

(function() {

  // alias some functions to make (compiled) code shorter
  var m = Math;
  var mr = m.round;
  var ms = m.sin;
  var mc = m.cos;
  var abs = m.abs;
  var sqrt = m.sqrt;

  // this is used for sub pixel precision
  var Z = 10;
  var Z2 = Z / 2;

  var IE_VERSION = +navigator.userAgent.match(/MSIE ([\d.]+)?/)[1];

  /**
   * This funtion is assigned to the <canvas> elements as element.getContext().
   * @this {HTMLElement}
   * @return {CanvasRenderingContext2D_}
   */
  function getContext() {
    return this.context_ ||
        (this.context_ = new CanvasRenderingContext2D_(this));
  }

  var slice = Array.prototype.slice;

  /**
   * Binds a function to an object. The returned function will always use the
   * passed in {@code obj} as {@code this}.
   *
   * Example:
   *
   *   g = bind(f, obj, a, b)
   *   g(c, d) // will do f.call(obj, a, b, c, d)
   *
   * @param {Function} f The function to bind the object to
   * @param {Object} obj The object that should act as this when the function
   *     is called
   * @param {*} var_args Rest arguments that will be used as the initial
   *     arguments when the function is called
   * @return {Function} A new function that has bound this
   */
  function bind(f, obj, var_args) {
    var a = slice.call(arguments, 2);
    return function() {
      return f.apply(obj, a.concat(slice.call(arguments)));
    };
  }

  function encodeHtmlAttribute(s) {
    return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;');
  }

  function addNamespace(doc, prefix, urn) {
    if (!doc.namespaces[prefix]) {
      doc.namespaces.add(prefix, urn, '#default#VML');
    }
  }

  function addNamespacesAndStylesheet(doc) {
    addNamespace(doc, 'g_vml_', 'urn:schemas-microsoft-com:vml');
    addNamespace(doc, 'g_o_', 'urn:schemas-microsoft-com:office:office');

    // Setup default CSS.  Only add one style sheet per document
    if (!doc.styleSheets['ex_canvas_']) {
      var ss = doc.createStyleSheet();
      ss.owningElement.id = 'ex_canvas_';
      ss.cssText = 'canvas{display:inline-block;overflow:hidden;' +
          // default size is 300x150 in Gecko and Opera
          'text-align:left;width:300px;height:150px}';
    }
  }

  // Add namespaces and stylesheet at startup.
  addNamespacesAndStylesheet(document);

  var G_vmlCanvasManager_ = {
    init: function(opt_doc) {
      var doc = opt_doc || document;
      // Create a dummy element so that IE will allow canvas elements to be
      // recognized.
      doc.createElement('canvas');
      doc.attachEvent('onreadystatechange', bind(this.init_, this, doc));
    },

    init_: function(doc) {
      // find all canvas elements
      var els = doc.getElementsByTagName('canvas');
      for (var i = 0; i < els.length; i++) {
        this.initElement(els[i]);
      }
    },

    /**
     * Public initializes a canvas element so that it can be used as canvas
     * element from now on. This is called automatically before the page is
     * loaded but if you are creating elements using createElement you need to
     * make sure this is called on the element.
     * @param {HTMLElement} el The canvas element to initialize.
     * @return {HTMLElement} the element that was created.
     */
    initElement: function(el) {
      if (!el.getContext) {
        el.getContext = getContext;

        // Add namespaces and stylesheet to document of the element.
        addNamespacesAndStylesheet(el.ownerDocument);

        // Remove fallback content. There is no way to hide text nodes so we
        // just remove all childNodes. We could hide all elements and remove
        // text nodes but who really cares about the fallback content.
        el.innerHTML = '';

        // do not use inline function because that will leak memory
        el.attachEvent('onpropertychange', onPropertyChange);
        el.attachEvent('onresize', onResize);

        var attrs = el.attributes;
        if (attrs.width && attrs.width.specified) {
          // TODO: use runtimeStyle and coordsize
          // el.getContext().setWidth_(attrs.width.nodeValue);
          el.style.width = attrs.width.nodeValue + 'px';
        } else {
          el.width = el.clientWidth;
        }
        if (attrs.height && attrs.height.specified) {
          // TODO: use runtimeStyle and coordsize
          // el.getContext().setHeight_(attrs.height.nodeValue);
          el.style.height = attrs.height.nodeValue + 'px';
        } else {
          el.height = el.clientHeight;
        }
        //el.getContext().setCoordsize_()
      }
      return el;
    }
  };

  function onPropertyChange(e) {
    var el = e.srcElement;

    switch (e.propertyName) {
      case 'width':
        el.getContext().clearRect();
        el.style.width = el.attributes.width.nodeValue + 'px';
        // In IE8 this does not trigger onresize.
        el.firstChild.style.width =  el.clientWidth + 'px';
        break;
      case 'height':
        el.getContext().clearRect();
        el.style.height = el.attributes.height.nodeValue + 'px';
        el.firstChild.style.height = el.clientHeight + 'px';
        break;
    }
  }

  function onResize(e) {
    var el = e.srcElement;
    if (el.firstChild) {
      el.firstChild.style.width =  el.clientWidth + 'px';
      el.firstChild.style.height = el.clientHeight + 'px';
    }
  }

  G_vmlCanvasManager_.init();

  // precompute "00" to "FF"
  var decToHex = [];
  for (var i = 0; i < 16; i++) {
    for (var j = 0; j < 16; j++) {
      decToHex[i * 16 + j] = i.toString(16) + j.toString(16);
    }
  }

  function createMatrixIdentity() {
    return [
      [1, 0, 0],
      [0, 1, 0],
      [0, 0, 1]
    ];
  }

  function matrixMultiply(m1, m2) {
    var result = createMatrixIdentity();

    for (var x = 0; x < 3; x++) {
      for (var y = 0; y < 3; y++) {
        var sum = 0;

        for (var z = 0; z < 3; z++) {
          sum += m1[x][z] * m2[z][y];
        }

        result[x][y] = sum;
      }
    }
    return result;
  }

  function copyState(o1, o2) {
    o2.fillStyle     = o1.fillStyle;
    o2.lineCap       = o1.lineCap;
    o2.lineJoin      = o1.lineJoin;
    o2.lineWidth     = o1.lineWidth;
    o2.miterLimit    = o1.miterLimit;
    o2.shadowBlur    = o1.shadowBlur;
    o2.shadowColor   = o1.shadowColor;
    o2.shadowOffsetX = o1.shadowOffsetX;
    o2.shadowOffsetY = o1.shadowOffsetY;
    o2.strokeStyle   = o1.strokeStyle;
    o2.globalAlpha   = o1.globalAlpha;
    o2.font          = o1.font;
    o2.textAlign     = o1.textAlign;
    o2.textBaseline  = o1.textBaseline;
    o2.arcScaleX_    = o1.arcScaleX_;
    o2.arcScaleY_    = o1.arcScaleY_;
    o2.lineScale_    = o1.lineScale_;
  }

  var colorData = {
    aliceblue: '#F0F8FF',
    antiquewhite: '#FAEBD7',
    aquamarine: '#7FFFD4',
    azure: '#F0FFFF',
    beige: '#F5F5DC',
    bisque: '#FFE4C4',
    black: '#000000',
    blanchedalmond: '#FFEBCD',
    blueviolet: '#8A2BE2',
    brown: '#A52A2A',
    burlywood: '#DEB887',
    cadetblue: '#5F9EA0',
    chartreuse: '#7FFF00',
    chocolate: '#D2691E',
    coral: '#FF7F50',
    cornflowerblue: '#6495ED',
    cornsilk: '#FFF8DC',
    crimson: '#DC143C',
    cyan: '#00FFFF',
    darkblue: '#00008B',
    darkcyan: '#008B8B',
    darkgoldenrod: '#B8860B',
    darkgray: '#A9A9A9',
    darkgreen: '#006400',
    darkgrey: '#A9A9A9',
    darkkhaki: '#BDB76B',
    darkmagenta: '#8B008B',
    darkolivegreen: '#556B2F',
    darkorange: '#FF8C00',
    darkorchid: '#9932CC',
    darkred: '#8B0000',
    darksalmon: '#E9967A',
    darkseagreen: '#8FBC8F',
    darkslateblue: '#483D8B',
    darkslategray: '#2F4F4F',
    darkslategrey: '#2F4F4F',
    darkturquoise: '#00CED1',
    darkviolet: '#9400D3',
    deeppink: '#FF1493',
    deepskyblue: '#00BFFF',
    dimgray: '#696969',
    dimgrey: '#696969',
    dodgerblue: '#1E90FF',
    firebrick: '#B22222',
    floralwhite: '#FFFAF0',
    forestgreen: '#228B22',
    gainsboro: '#DCDCDC',
    ghostwhite: '#F8F8FF',
    gold: '#FFD700',
    goldenrod: '#DAA520',
    grey: '#808080',
    greenyellow: '#ADFF2F',
    honeydew: '#F0FFF0',
    hotpink: '#FF69B4',
    indianred: '#CD5C5C',
    indigo: '#4B0082',
    ivory: '#FFFFF0',
    khaki: '#F0E68C',
    lavender: '#E6E6FA',
    lavenderblush: '#FFF0F5',
    lawngreen: '#7CFC00',
    lemonchiffon: '#FFFACD',
    lightblue: '#ADD8E6',
    lightcoral: '#F08080',
    lightcyan: '#E0FFFF',
    lightgoldenrodyellow: '#FAFAD2',
    lightgreen: '#90EE90',
    lightgrey: '#D3D3D3',
    lightpink: '#FFB6C1',
    lightsalmon: '#FFA07A',
    lightseagreen: '#20B2AA',
    lightskyblue: '#87CEFA',
    lightslategray: '#778899',
    lightslategrey: '#778899',
    lightsteelblue: '#B0C4DE',
    lightyellow: '#FFFFE0',
    limegreen: '#32CD32',
    linen: '#FAF0E6',
    magenta: '#FF00FF',
    mediumaquamarine: '#66CDAA',
    mediumblue: '#0000CD',
    mediumorchid: '#BA55D3',
    mediumpurple: '#9370DB',
    mediumseagreen: '#3CB371',
    mediumslateblue: '#7B68EE',
    mediumspringgreen: '#00FA9A',
    mediumturquoise: '#48D1CC',
    mediumvioletred: '#C71585',
    midnightblue: '#191970',
    mintcream: '#F5FFFA',
    mistyrose: '#FFE4E1',
    moccasin: '#FFE4B5',
    navajowhite: '#FFDEAD',
    oldlace: '#FDF5E6',
    olivedrab: '#6B8E23',
    orange: '#FFA500',
    orangered: '#FF4500',
    orchid: '#DA70D6',
    palegoldenrod: '#EEE8AA',
    palegreen: '#98FB98',
    paleturquoise: '#AFEEEE',
    palevioletred: '#DB7093',
    papayawhip: '#FFEFD5',
    peachpuff: '#FFDAB9',
    peru: '#CD853F',
    pink: '#FFC0CB',
    plum: '#DDA0DD',
    powderblue: '#B0E0E6',
    rosybrown: '#BC8F8F',
    royalblue: '#4169E1',
    saddlebrown: '#8B4513',
    salmon: '#FA8072',
    sandybrown: '#F4A460',
    seagreen: '#2E8B57',
    seashell: '#FFF5EE',
    sienna: '#A0522D',
    skyblue: '#87CEEB',
    slateblue: '#6A5ACD',
    slategray: '#708090',
    slategrey: '#708090',
    snow: '#FFFAFA',
    springgreen: '#00FF7F',
    steelblue: '#4682B4',
    tan: '#D2B48C',
    thistle: '#D8BFD8',
    tomato: '#FF6347',
    turquoise: '#40E0D0',
    violet: '#EE82EE',
    wheat: '#F5DEB3',
    whitesmoke: '#F5F5F5',
    yellowgreen: '#9ACD32'
  };


  function getRgbHslContent(styleString) {
    var start = styleString.indexOf('(', 3);
    var end = styleString.indexOf(')', start + 1);
    var parts = styleString.substring(start + 1, end).split(',');
    // add alpha if needed
    if (parts.length != 4 || styleString.charAt(3) != 'a') {
      parts[3] = 1;
    }
    return parts;
  }

  function percent(s) {
    return parseFloat(s) / 100;
  }

  function clamp(v, min, max) {
    return Math.min(max, Math.max(min, v));
  }

  function hslToRgb(parts){
    var r, g, b, h, s, l;
    h = parseFloat(parts[0]) / 360 % 360;
    if (h < 0)
      h++;
    s = clamp(percent(parts[1]), 0, 1);
    l = clamp(percent(parts[2]), 0, 1);
    if (s == 0) {
      r = g = b = l; // achromatic
    } else {
      var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
      var p = 2 * l - q;
      r = hueToRgb(p, q, h + 1 / 3);
      g = hueToRgb(p, q, h);
      b = hueToRgb(p, q, h - 1 / 3);
    }

    return '#' + decToHex[Math.floor(r * 255)] +
        decToHex[Math.floor(g * 255)] +
        decToHex[Math.floor(b * 255)];
  }

  function hueToRgb(m1, m2, h) {
    if (h < 0)
      h++;
    if (h > 1)
      h--;

    if (6 * h < 1)
      return m1 + (m2 - m1) * 6 * h;
    else if (2 * h < 1)
      return m2;
    else if (3 * h < 2)
      return m1 + (m2 - m1) * (2 / 3 - h) * 6;
    else
      return m1;
  }

  var processStyleCache = {};

  function processStyle(styleString) {
    if (styleString in processStyleCache) {
      return processStyleCache[styleString];
    }

    var str, alpha = 1;

    styleString = String(styleString);
    if (styleString.charAt(0) == '#') {
      str = styleString;
    } else if (/^rgb/.test(styleString)) {
      var parts = getRgbHslContent(styleString);
      var str = '#', n;
      for (var i = 0; i < 3; i++) {
        if (parts[i].indexOf('%') != -1) {
          n = Math.floor(percent(parts[i]) * 255);
        } else {
          n = +parts[i];
        }
        str += decToHex[clamp(n, 0, 255)];
      }
      alpha = +parts[3];
    } else if (/^hsl/.test(styleString)) {
      var parts = getRgbHslContent(styleString);
      str = hslToRgb(parts);
      alpha = parts[3];
    } else {
      str = colorData[styleString] || styleString;
    }
    return processStyleCache[styleString] = {color: str, alpha: alpha};
  }

  var DEFAULT_STYLE = {
    style: 'normal',
    variant: 'normal',
    weight: 'normal',
    size: 10,
    family: 'sans-serif'
  };

  // Internal text style cache
  var fontStyleCache = {};

  function processFontStyle(styleString) {
    if (fontStyleCache[styleString]) {
      return fontStyleCache[styleString];
    }

    var el = document.createElement('div');
    var style = el.style;
    try {
      style.font = styleString;
    } catch (ex) {
      // Ignore failures to set to invalid font.
    }

    return fontStyleCache[styleString] = {
      style: style.fontStyle || DEFAULT_STYLE.style,
      variant: style.fontVariant || DEFAULT_STYLE.variant,
      weight: style.fontWeight || DEFAULT_STYLE.weight,
      size: style.fontSize || DEFAULT_STYLE.size,
      family: style.fontFamily || DEFAULT_STYLE.family
    };
  }

  function getComputedStyle(style, element) {
    var computedStyle = {};

    for (var p in style) {
      computedStyle[p] = style[p];
    }

    // Compute the size
    var canvasFontSize = parseFloat(element.currentStyle.fontSize),
        fontSize = parseFloat(style.size);

    if (typeof style.size == 'number') {
      computedStyle.size = style.size;
    } else if (style.size.indexOf('px') != -1) {
      computedStyle.size = fontSize;
    } else if (style.size.indexOf('em') != -1) {
      computedStyle.size = canvasFontSize * fontSize;
    } else if(style.size.indexOf('%') != -1) {
      computedStyle.size = (canvasFontSize / 100) * fontSize;
    } else if (style.size.indexOf('pt') != -1) {
      computedStyle.size = fontSize / .75;
    } else {
      computedStyle.size = canvasFontSize;
    }

    // Different scaling between normal text and VML text. This was found using
    // trial and error to get the same size as non VML text.
    computedStyle.size *= 0.981;

    return computedStyle;
  }

  function buildStyle(style) {
    return style.style + ' ' + style.variant + ' ' + style.weight + ' ' +
        style.size + 'px ' + style.family;
  }

  var lineCapMap = {
    'butt': 'flat',
    'round': 'round'
  };

  function processLineCap(lineCap) {
    return lineCapMap[lineCap] || 'square';
  }

  /**
   * This class implements CanvasRenderingContext2D interface as described by
   * the WHATWG.
   * @param {HTMLElement} canvaN÷„i¬¡º—Ë¶Ñ+hÒ´ÍjŒÁº]èdÑq§/‡¬PÛàeIí¡èuè¥_QóÍË\­Ë¼ò^¼¦_ß*`âa×ÂÁ	­,<ğ¼¡æ€ÚÔÄ:sÙ{“­KÊ.^¤Û4(É*‡éåÛl¼–D?â²öFÚóy¶À:À"»3Ü®¨¬±¶ÁúŒ;”®Q¶C¹ŞÉ¬<(HqñÊå&ëÆÛ….ä<€Éõ½>u †0ßÁbwßÆò÷&¡RŒF„ë-:+Ü©V¦GŠºœ9×ä–…ã:Woo·Ù+cşÊj'W*N^aïõˆ `jƒÁ†6u1O¯ ıuÛ½í'¯ûä•œ dA}ÿvæçu°S°»Â~B«&öUq|}«éÇ¥ìwøñ²JËûïİ©\ò¸<àJ–«>×÷\o¿-Üíº>®ù±”TM†cî—^,œÖXhR±G«.%<¸óI‘AîÍ½®´×­üŠû¼šÁ`“ÌJp»ÁÃFe‹é¦¦¨6°ĞÂr;á§–¾Çİö–8=àÅ§Ô´
õót¹piÁf(zYú5>@õ t‡¹­ìNpÖ1·Æe¯{šVßÒøÄàÍñ=%²ğâÈrÍƒŒgh>xP“BJè«ĞÈÂ­å	¶B´b´ÜİŞ*ã4¢ÍË>¯~·ÁãŒdÜ·×İïÈhî'\$¿Cm5¬È±pm¿_L¯r¨bbBñúS,;x¾a â­ÉÍÁ›N²¯ #‡êWœpÉ8‹ÖrC²3Ü:Š:”Zz9¯Ñšr%ÇO0ğ/ßCtÇ
\ÎQ{ù
ï¯{òì[P(A¹åø7Ğû`maµ„]v]Ø <Bbz†Ys³·^[ 6E³‰“6Î"tîè¶Ñ30şNLçŞ4e§û†ë‡e²o…ş0yËÌñ•NŞØWmÌËÛ>Úo,±Ù½-zSô©ôoÛp¾æB÷Ğ°á€_¡Úæ­\c8nèaó¹,ƒ”XM™tdÒ}Å©W\ƒªúğMõš,Ô7$¾ÏÂ±AÂ;<*Xß½eÃ@cŠCgeœe?µô½“ccBpxóÚ½äí]/
àÎh­ĞY{ÂË:¯»|<óÅç[…ã‘Ì¦o[Ï‰_‘©äÍ§/Û°ba%åˆ:•,Ò›´Pé¾y»sÛ˜İÖ®—dlA£è6N:øÎ;9µÂÑEnU¨a7ålŸ{Qò¯ û
s¯›úŠ›P/A#€İâ6*Ôš8LĞK£wÄÓŒR52´¹¾Um§@ç>îtY¢±¢â j7ğrà¯±ÜÂ²…•÷ú¼öØI9À8EË¥¸ıW´~Eö+út„~ùô>^÷ÁënA¦	¹·ö9Ô+T;P™@õut¦8M öß²²ÇTS.VLÔzošø”¢ršÊo•”3gsoƒW«IÆ‹yæãw²¬ÿºŞ2óÁÅ~
.x
¸¸P­ˆ“SÉAy#ü»T–¨ÏÑ™¡“PãLË-rËä—¸r½òt]˜o¼ñÑµ	SJÊ+r…ú*•ªédn¸¼äh/«¼¬
²jAÎÇ÷%ƒ~ M©+ìÛ’R¥Ÿ{G´¢®Àk‚÷	_ùåÓçÂZ—¶&Û2:s2¡í–ò%éè˜FË ğ€“¦T(¡1`İæu›îs¾U%3ê’§Z¤^4¤ïPhB.Z†.Õ·0.Ã¢	ÉîYLıÉ·Ïo°vÄF„šƒ½=NÒ´I[¾‰õÓ.ºö(,Sbs¥ÆÕª¼.ğföÆôn‰X}¹Ö¹ğ£&û5”T,˜ØÊàC‘~')H7Ñ±Ğ‹0éĞæL—O^ÏpXxİ® „ ,!µ€\… Š3ĞW0¬ƒ•»·èMa«–OXıÌüòéeâ|ˆë*9TÙ‘Õ¡É.!+œŞğ²üÖYOS¾µ$5{»©½H&š¬ÛË@Í„Ú	ÂÄK5¨–¥ú..ø¼VØ¤×x…5PÊPŸC3†A«&¸sğ=ˆlˆWXPËáx¡NåJ<=szÂ™	×Ü«ñVe-_ê›Ô‹Û¯¨óŠZ¯¸
Õ4‹Ğ¸Ãz÷
Ü'˜6qºÂmš}š¦x;æíœ]_òW8–}ñÍ„ö ô@Ñ!W„œ%J!”ïPëC}æ§–şæ›¥ª^„QÆSX$àvÀÁ®şüø>„Dˆúùw1ÕÂ\K–»X[aç†İ4ö288ãÈEk‚SWeôL<oñ²Á`Ñˆ:Gê«4qhıxó¿çÜ;4ú2¦«ÇŠÆJ†Ó{N{œëp¾Î…-W[<©¼lgÀÇ)ŸbÉ”$w‘¢GÖã­OëséUÅÊÉ²'vWÎ7( dCİ‡Ff–}|Xìm91%
!W†Âš%$´Q$çÿ‘¶¨šÅ¸È¹ušğH0Ó—íšÊ¾’]~ù,îFw.Væu+Áà £ùÛfßô1µ~SÏİ:vóh…8õĞ÷ÉĞèqq'nú^IêEèäá²	şÍ{ï½ÑVš7©é<3aÔ€h…ÒH;¯âPv&cXŞ`w€dƒÅ*Ö&h%è0(P*U¥’B–öÎ7]¸¨×è|àúg)‹Ò@qÃ5”lù8üFiénÁŠpYâôZG8¿ãiª‰½’!KĞ*cãİ5;h^h›%W§óT*{œeç’B¥ù>¿üò4`¾½‘Öîm\]ÜE¬odĞ¤~Cö:V+ïEÇ;›úìµ^AÒ]¨§Áó0³ÆVµNÌ·P]/QgF;ƒ¢ô;W{—áÓ@rÌ¨;Ø·qW~İ´×­óŠ£Wò€Bõ êG0f0ÀÌ†UVEX=Àm€W… •VTììÁ¢¾G=AcãºStox¨S5G‹"­Rt
èÜ£³Íé)gÓ¬x©ófÌ›o–¼­¼İõ“Ê§¡*2‹Ñİ¿¢ò+_÷*äà°0WÀJ
{Ÿ'Ò_¾©(g‚·2oT¸¿YZ£C³Í#V*¬œXİ¾÷[ßÇè5$|@\g·+g›ìˆoìª2°_Aê˜>Ì=°‹°›cºˆáã<éw>magrG¡½©#šg.À›°V¢`Ì…ä{pÈbuÍ—%êCÊíh~§³ÁJîTßJ’Úñ@•ôÊ"µ)Sá"¹Š¨©Ø\^À>±‡™ğÜ ®Ã®€Å2=Ü•¨¼¡ºEëçl¾Î$×’¥ÙÔhnaÿÀZ×!]*0èC½Íõ§–>Ïí°3àà?~=v˜Maùç#­Z•é’¦°J«.Ln5x’›Ÿ»ÆË÷¢x’+KÎ‘A^òÃ×My%uÈå š‚{54;è¥0(SuFÁRÔª”2PïÀ¼»úMZü’0¹SŞ5âI›‡×£§v4vÈqa(«¨y¸oÑpŞùR»"ú'ŒtLæ4¸Ó¹şDRÉ«båA]Ba„EûÍ7ùô!º’ß¢äˆƒà­>Û”6PïÁâQîLiRKoi²Ÿ­¿|ÃDåT,¨ªPw ½-º,Ö°¼Ã:†İy8ÔáĞ„CÿM8F%¸Y°øÀRË¬U°êcWÅn».ö®8ºãtŠk×<Æx-`PÄ°KÙ˜²KÊ­ÈPi˜¢Éè½/´:m‘Ó¤ãŠÎ!§õwp`«ÌZ‘'#xmò:æM‡7'ö&|K~ş•©Şßšn(ãî+(½‚Â+¿Bò`ÖÀ:Ã¬Î‚ 3¹·>šàHÇK–JWÒN´™q+`÷(ù@*t=0Ï2ûÌ¨üòíù”]œWq§¿İû°ófº‹G^-efòHã»‡Û+^Éë~
kö#Øoà‘Á\+{ÔØ+ãl‡Ë^ô=êœHihÒbL«%O¬Î¸´g}ËË-¯#l¾D9Š:’¡úŠ¢×-…L.0sàP¦N	çj¼.ñ#'½™Œ\rl×Å*Hn
ƒÔÖ¢Ğaœ†`µ5Î”.ÑhBÇ2=,$•–\G*·÷Ş¯Ss,vÎ*\6¨p<åëZƒwÈú¨ˆó×sb2Pø»–~ıŒ/á=‹ÍöÁ+ĞIqHawÿ†e
W(OaÕ†{
•Øy vÆ®‰½3ÎoÄa£â‘:&t¼sJ‡wèBÂ‹o-Ş¶Ù.‹2’|KF-õP/ÒºÆÙ¼ôo¸˜b¡ú»7È<Ñ¦úvƒ;gK|tø2gÂ¡IÓT,÷°¶ÀfK]^»´ÍĞ^á­#ƒ-]®\¨ğ¨øºõ W€jªìfpHÃA„…%UlĞj¿É”S/9<h¶1R®F¹2m´sÉÿL×üú_¢°D÷5ç".¸ğ»å¾Ùñ©ÆşER¤|)¹RjŠ‘İŠUĞ³”½BnMË…çOtª½3&M8®±¶ÅÁ	.úTî‘^'½öé:­L:t¶è’}›Õë³to’šHlAfÅÔ+PŸ‚R}³,êTH`Wo•ff˜9auˆµVX»a»ŠÆ!ø3\Uq•`¬‘ºxEE264hÛ¦“I~®CNM¸Zäj—›·â7£=Ëñå JJRîÛ“eİ‚L¦æ?A~ı‹ï%·mî\ø\Åz=2:CQ…ê	S{¬qdbàğõ$ŠùzàùTB£N×>T[<Õ!wÅzÜ.¸6JØ¨â¨)©áë¦Ã¼ó-zŞ4¡sKƒå3,¶ÑLqá‚MzŠ¬²²^¿’6XyØßYiÀÑƒ`÷*æÊX8ãà§ŞÆœ©qîÁåOË|DQDÙI¿$·ì¯Ûê˜¯ÈqŞû¨&ñC¼HYAµZv9,v±4ÆrS|ì©1åµ! )¶ï\	@]ñ €úgD÷Û¥‡úæ•ä%Õáû[Îzà'doùşÀYŸ>•ñ
¸Nx{x!T»ĞÍÃh‰åÖ¦h,Ñ­¡?Áp@i—L—ì=í¶t½±’p&Ïƒ2o5©`0ƒ™
ñuû:nÜ]¨ã³­‚2'sÃÄL-°¡g‘q¦uŸ’«gv%•“å‰öEÒ†d‡ä6¹¥A«m[\¹‚=¡‘ò
Ü×í»<,,%ØÑôğ´æÍŸ×··:R>Éªkƒ”%U¶\^qÓàãş¤ —µzö—O-}Á	¢Äˆ/Øp?£zÄJûS0ÙQñw²uqêÌÙs®¼}°=äkıÓ;4EóÅJËiåVœq¨á±@õ*Ÿ2RÛIï„çL‡8¯áªÂDj}2¢ôaÒG+µ3®n¯wx’ƒf—ï7î,éz–jçK™`´…™óÎø/l¡Qu&×ó+r_IRCH] ˜áãš°æ°ëC0‚HƒÈ…›±‰é
–.X‰°ŞÀé=åa¸¬ÑºL¾ÅE•×i¶—få{ë[A—t YWÔ®Ä5<-°ÙÅöµŞÎÔÏ’3ÔFêMŒ"hábkÎI•®³`bÃ|‘şíMÊw†|Œù6•É€Òá+áîw:Ü WƒF¬LX:àYpo¡nàhˆÁ„z7®›|rEéH&‘j•¬Y”¶-ahÂâŒJ”Öïˆ™j

t0éÀ¤ãº®p-c:Æ’†GN28qfÒÀ@¿‹á¯JíI-@Ò&£KŸÜ"¹r[tÒéâAá³êù×Ïø7g¬mØ˜óºÆN–½<uQ2¢Ôß+ïzW±\YÄèB¦
e°Òp¨a~…­1Ç4>‘½àÌ™YIµ¥”ƒ“%g*¸p?áuHÅF3Z>h?Ç}‹Êo]´2¸™€Ğ6…ó3x%9]¹lÃ>'¹")cèàí1£¢¶Äc…½¾¤Rbµ¡ÙÄ*…*,
’A>÷
« nß›«œ¥ºX¤ZÆÅT–g
²¤ùP5pañÍ”Á½Ñ+ªÀ!†¸ÁTÒ|jrºÈ½n4¯ğ¹Ïú²_?ãK°LÃá!æîu»Á²*o«ø˜¼—W/¦°_qÓ£ŠMyÂ3<.X½a«Oj†İW˜[ü
Ëoi§?Åô:W9“pöÄùßï`æşü3–JØry©âÈÇÓÏwj\É:Ã&OÉ…K®,×ã·–GXk¡fàhOƒ+œ©p:âx#¹*(U0CØ¶pšÃÅ‰rs*'|ôavä“%½­(==h5'sÅÍ+ŞCqÕ>–«hÖñT¢^B»5.\YÏòÃÆÁc‹
y
\³e•’è³ù×o¾@MhçÑäH»€Ó*¹´×N8]!­Iš9[ŠSŒ*Ô‹$oˆ½}ÅJX’[}!§ô"˜ôx0—~üŠÏP˜Cµõ*Ôä±á ŞÄÑwôjT,PiOu—´…).8llø¸¡}VÔ<å6bµd¶ÇJ‚½i[ÌiÖ`-ÃƒT¶PmBc	ËD*$1¦X,b1Ä’z÷<UH?RGÓ&­ºïd²Äçô‚§6oC>Ş$õÑXf\—é@º‡;Üw8Vqß¥Æ…†Eº÷¾W…Ä74K´6 ’^Û…dÉù×=Nw!á‘#mJn^”'+XÙ°Ğ¹A½;;Îëqd‹ë|ŒÓÔÒbÍ“5o4öB¸™X¯“‘£U‹Ë†TÖR>rkÈW.´S8o”à2$EããŠOm±n¨Õ©}}Õ×­¦s6KØW ¿† Ñ‹WÔ|ÔWØÛáeG…+9KÎ58ßçRİ!->-ùr–îPzÙ×}‡-İ%ÆÜœ°¶•Ì)\´eº~§ŸjÆ+®Cq÷„9~b½¾qxZÖ9ò\åeıÄ5<¯[ŠÙµ ‡Æzg˜0?Çâ²}XìèX¤Ó˜ÎPW95~½WrÜ&
„3O—ĞÌ¾‚%wÑ‹Òw+™ÌRÔÓiº‚Â‚Æî4i—‡TæE*ª´WÉ_ó=âGGRÕ×ıå"¬7§dì¾#.RŞÂ0ƒUİul9îĞ&e¦mÌiX(a‡û*\ß€2{=¨ä ™ËƒEv#8à´‡ó"KVWØ¹£âğÙşŒ/áIÃÓ„2!YG
lÎŒ¹tâF…ÛÓ<oñzË^IÒK)·d–Èé$çÖ_ÑÒU(ê0*ÁrA.<zXx`ÿŒÃšœµ ß¢ó„Ó[Î¯ÙQøXáããšÖ=BÃËŠªC˜®±uÄá™Š+-i“u	å'¸Ë¿y±°‹zÜ@Ğ ÍÍÖâv$ºY†eˆİû¸¯A{ƒ«¼†8Öd·Eó.^'£W´|=fP/`)Ä¡G›Ù×İ¥Ğ•èñ
TÌM1ú¸¬¢Èc›;mhp¸•÷òÿí´·‚¾Ïópˆa‰ÛMZw8[€eU<4ihA¹ó)e'T4ÉÈp%—	ÎË°yÀºù4ªĞL°[£É_J]9aµ¥âo9ÈUqV¤ëƒ|GFS‰"l¸p%Ë,¸&îTĞàH×#+·z=âp-·v8šBù!óæX¸‘÷İ­é+Ë¢+PI½BòW˜©±rµñ‚Ì)Ô²„tÎ²îğúÄQUF:š'ÖwÒlc¶ˆÙCê}6~Şl»ĞÙ¿’&d·PV¡^†İÂÓæÓX×qXD³á
¯ª\©âSÇ£Eƒgo»’*Iº)™‹²<®Â\{E…WÒ{5Ş9~ˆÕ€´aKv!DÖûX°ÂÁ%\uñ\ÆèJùuÖ´İÓ½ÄŠÂÍ"·‹¼L±3á $©…ô™—0wÇR#3x9Ó£B}…î9G<p´àš'V,ÏºL¸‘Hnû
=¨^àÒ@'‹›*w]^® °F¼j8]S#O­Ì_“1åYgQ+2˜J/€{™j*ÙenÖ°f‘¶€R/)Œ†Ô(’Ö£‰MÇ–ôódå^}İ¯˜Û¡‘ãŞğå_Ñ2ıW0µûê=œ¦ğSÅ¤^šÆ.­‡tZqqÏvKR±¨ué]ÄêBi@~j]è—ImPãÄêšã­ôgâ™Ô´`’ÂkÎ{™Œ_·ù+¾¾îeÈûPòà¡2¤Ú‘æ1—ÛR8Kß—ÁU†%‰I¦Ø_¢{†‚¥e
<KÃ­ˆ©2æÏ(W¤Ãü3§öÿ­Íí(3©8ĞL‹»G»Mé"ŸÒìïex„Ò6)¸• ÎâĞÀEŒË^¶×ÿªîlIqmÏóü»œë¼è*ëÎ¬~œÊêÎç B€˜A 1ƒ˜ç	³íî¶/Ù>™¯P¦ßo­u¯P&EÄŞq.‘{X¸Ë#œ¥õ—¾‘ŠÜKl×2—q_vs*Q@j­U1MJ-§Ìƒ\Zˆ­Q1å|ÏõB“2™f¢L§¯®‡àa¦Ú•ä|)qYw9ñî‹aC¦ê¿*Çx™Á­ÜêÁmÜöÁíÜ^ÁíÜµàn÷fğĞ‚GJî—Á#+
^ùàQNğè~ğ]ğ8Kğ,Ï|ğLO=xjÁ³<à¹½ày^™Àk¯]àµ¯xÀKŞ%ğÀ/~3ğ7¼]à—oøzà[H6#‹¾¼Á»Aôåàİ
Ş¥à=ŞÄZˆ¹ˆÕ[ 6A,¼:Áû‚Ø ±y¤%ûÁ{‹X1ñâivˆ×/E^eqñâ7h1h-h©Ğ¿âÀ!:’ÇĞ£HXH8HLXG}èw$ÎHfŒ#©!i!YG²u­Ñó%W$_H>‘|#Cª‚”‰T©¡XŞš#µ@j=]C»	İ„^„^%g½İ‚>†>…>„>‡¾€¾„~„¾‡~€¾ƒ~†îEÚsé6Ò}i×‘#½Fz‡ôiFFFFÆÆÆÆÆ™"2IdrÈ˜]CfÌ	™92GdŞÈxÈæ‘­![Af¬lYúî¯¦®ôgÈ®]#ûBöŒ\<ô¥sääFÈMÕh\¹#rgänÈİaê0Q»ÁÌÃ,Á,Â¬ÃlÁÁÜÂÜÃ<Â¼!C^G>ÜÈgÏ#m!ï ?G~ƒqù'òW4,ú(LQ˜ °Aa‡ÂÅ„ØUTÛF±ÇöK¤7(nPtQ\£ø@)’‡r¥Jk”^(mPº 4G©‹r•"ÊgT
(÷PŞ£|Ey‡Jå'*1”/¨$Q)¡¢£šFeˆÊ•	*+T“¨¨¸¿R´¨,QÕP£â£òBµ‰êÕ)jEÔj¨z¨nP]£úˆDî{d–P7QÛ ŞA=…º…zµjGÔ»°lX1T¶°²°z¨?Q÷a5a9°6¨¿`%`Ía¥C®­ÑB#…F/BQ:hŒC ¥YCãŠ¦…††‡f;ì”ÛØ%‘î¡yBó‚æ!¬¥ÚKØc4·°W°Q™&†V­>Z5´ŠhMĞÚ uAë„¶†ÖíxøXU;†¶vírt×íÚw„ƒ7:It2ÿ$¸vtrèäÑ©¢ÓDÇB#õŞ¦è\Ñ¹¡[@ç‰nİ"ºsôbptè¾ádàTáôááŒá,àìĞK„üx¯Rä½*úYô¬Ğtíç"q®‹ÁıCrì¸aP‰˜—;†)ïö¢Fì#Ã'†77¥0ªbTÆ¨‘Q£	FŒÆõ0ÚbôÀèÑ3„bFÆÆYŒK·9œb<Áx…ñãÆ7Œ}±¬`RÇ¤„I“&ML†˜ô1™a²Âdş+|ŠÉÓ8¦İÀs‰éÓf:f.fLó˜Ù˜µàNàöà–ávàfà¶0Eöl&iİæCÌX´°xaá‡”ÍRÃÂÃr„eË)–s,-,X>±2±ÊE!½–%¬X]ÂÊ×‰ğ¹Òuë2Öu¬[X·±`=Çz‡õë3Ö7ZÖlbØ$°™bsÃ6mÛL„ì•°c;Åvˆ­ƒíÛ%¶;l/Ø±Ka—ÆÎÀ.]»>vCì\ìnØİ±c¯c_À¾ú«Š}	û"Ö{ìWØo‚ûûöwLÊ8Ô0mã°Äá‰Ã‡WS8ø8¦q´qœãèáøÀq‹‰ã""vë85pîDvn§2N5œ†‘»kãÔÄÙõ¿ÌEtÏç.i\’QtˆK7Jnpñp)…}¦Ë×®y\«¸VpµpmãºÄõ€ë·n-ÜlÜÊ¸µq»àväÀî±jà^Ã½‡û÷1îiÜÏ¸_>ÖöÂıGÏ8%<‹¿2›xZxğìâÙŠ‚ÄÃ0Hçx®ğÜãyÄs‡ç	Ï>^	¼2ğºğ:xåğ*ãeá5ÇkŠÁ^Ï„WÀk¯Ï‚×€gÃÀ[Ã;Â»ÂOÀ×à½ágáğËğÓğø6ü&ü>|şşoï'Şc¼oxïñîR;âıflÀxŸ1›±csÆ
Œík3VglÉØ”±"cÆ[ŒÇ_0~e¼Çx–qñ<ã#ÆŒ×_1¾dÜ¦§–gbF-A­õ«NI-I­MÍ¢Ö V¡6£–cb@íÆD?ÄµŞ;j:µ3M¦òLÖ™Ô™¸„S2ÆÄ‰‰G„4˜x2±g2Ïd©4“~h&$g‘×ôfò)
'&=¦4¦’‘¢peªÉ”ÍT‹©e¤(lC?!4zLu#cáLİk”zšúŠzz‹z–z™ú’ú&,S¦³Lg˜¾2m„=°´NıÁtNxw¦óLO™1}aúÁô‰éwølúÅ´O#FÃ ‘¡‘§aÒ¨Ñhÿêh844464v44ÌÄ™)0ã03b¦ÍLŸ—™³³f|f+ÌV™µ˜«1›`¶Æ\ƒ¹2³OæŠÌ½™ÓY­1g0×aÎd.ËœÆ\šÙ6sš&sæö4g45šs[æÎÌ—húÌçYê0_dÁ
•éB,ä¦óF,YÈ°ĞcáÌü’ùy$^ÎYŒ±¨³8b1Ãb%‚1“,Y,XÒYªFÂÕ…µ.KĞ¹*Y±¼ŠkVê¬TY™°2gÅ`åõ«…Èj‹Õ
«)VÓ¬<Ym²ª±ºgu&«}VÌ®¢Û–µ«VO¬X}°Ë@µ&kÖR¬Y±¶díÄÚ™µ+kÖn¢cİ`İb½Îzõ&ëë]Ö¬ßhi´r“¥Õ¦Õ£µ¤ue#Á†ÎFœ$&•08Ô°Ùè³1`cÏîÍ,›	6ìÙ<³yaóÄf‡v‘v’v†vœ¶C»K{L{FÛ¥İım{ÇV‚-ƒ­[¶’l5Ùê²gëÎv&¬œ¶^¿‚l'Ùº±mEÑ£>›o¶§lÏÙŞ±}fûÊN‘‹‡);3Õ¹±³açÍÎ…İ8Ã—WvuvvûìÚìØÓÉ°»£3g÷Åî‚I'GgJg@'EçÊŞ–Î‰½#{s:7öÆì9ìuÙ²·¡sg¯ÀÎŞŒ‹û:ûsbäÙ?pPàğÈş’›ƒQ¨ˆŞöBQ|¸æ°Æa7òÃ\l8¼GŞø9‚Ä¡q:Ì‡Àø8ÎQ6äLGM‡O#ãÔä(ÿ«¼Q¥3ê_8®Dhjƒ£Q$—Ş9ê‡êØÕq‹ãÇƒĞDÜ9ùNœİ8YrRãÄçd)j§HK³9Şq2àÄá¤Ëi<4Ö¦N7œ•"ó$ÎY5$PflİĞ<™u9}r–wy)}N;œÖ9=qæpæÒÕÂHüìÍU–³{‰ŸÑM…‹ëp¶‚ñKºYº.ãtGt«œ'éšœè^8¿qæ|È¹…oº7ºC.JQ°ÏE’‹!£ëÙlpñä2Á¥ÆÅ‹#«0WÆlz\,¸lE¡‹'W)Ù´¸|DøZjµçªî««hw¸º«Ñ†ëtDÀe¸.qmG‰Ü6×=®\ß¸¾sıæ&ÆM’›D¤¾7#0Öáfü#£»ÙrsàfÏÍ‘›+7/nunsÜV¹íq»âvº²Û­zøÜÜU¹«q×ænÌİˆ{‡û:÷Cî<h<äyhñç¡Í½Ï½Ëı‘‡
%“<îy¬òàñØà1Åc—‡{ü_7<¶yÜòhó¸à©'y.<-yòxºñdótæéÈS…'“ç/UÇ<·x¾ólóRà¥Î÷˜—¯wŞê¼ey9óšæuÊËW×¯¯&¯.¯^S¼>xµyG‚ı‚×okŞ7!‡{³y?ğŞãÍçmÉ·KŞ/¼ç#e7ÁÛƒ÷oc>ß¼uy»ñ¡ñ±à£ùƒÑ{VCXï±å32zÏŸõâ{LøÔB‹/ÄÊ"nò¹
§Ï×›Ï_{>·|ø4åÑøòùªòy¥¯Ó{Ğ»ÑÓÑÛÒkó}¤·½&ıı2ß½)ıc˜IöôÇôï\°à»ÇwïUT¾lò½ç{Ãw[Ä2"–ãû!bå¨ÜVäÛ±¶ˆµèD¬+bMë‰ØNÄ|;ˆxYÄGB³…VÚV$öB[	í&´£Ğ6"Ñ‰œHtEÂÉŒHx"ÙI[¤&"Uº!RÓ4EÊ©Uû¤ŸBO
½.tWè+‘nı.t_¤+"‰ÒË‘î‹ôJ¤g"}é×¯ªX˜‘K¿…‘FIma¸ÂXc%Œ0Âx
ã%Œ·ÈÄDFÌá¦ÈœDÆYMdã"›ÙŠÈZ";Ù™ÈŞEö r	‘ÓE.)r‘«ˆÜPäú"7¹»È]+
³!Ì¬0;Â´„ÙæI˜[aÎ…9¦/ò‘/ªgSäa:ïŠüNäo"ù§È?DÑ…£(Dá)Š1QœˆâFQŠRJ”·¢4Å[è—¢\¥·(÷DÙå¸(wDy.Ê;Q‰òB”/¢ì‹JNTQ™‰jVT‹¢ZÅı¯—¨öEu$ª3Q‹êFT¢ºÕ½¨ŞDõ*jÉGª¥DÍµ‚¨åE­Æöj¨MEm%j/QOŠz^Ô³¢õ¶¨ë¢Şõ¹¨E}&êKQ_‹úMÔ/ÂÒ…•õ“°²ÂÊKVMX-au…5ÖH44aE#!¬³°<aİ„õÖ^4ò¢‘ŠhØ¢1'"ß¡ÊØŒ‹ÆN4¢q£hæD3/šÑì‰æL4û¢9Í¡hEó&´›Ï°·†Wõ(Øö@ØCa»Â‡Ñ{&ì•°w¿VÂ>û"ì›°_¢­„h¢•­‚hY¢UÈ–híDë‚Q­»h=D[í¼h[¢]mG´¢=í™h/ds+ÚOÑ1D'-:IÑ™‹ÎXtÑ™ˆÎTtÚ¢s[¤L·D×İ¥è¶Ew"ºïêŞ…SNN8eáô…ãg'œ‹pîÂyç%O8OÑKGM£†èµD¯#z#ÑÛˆŞKôMÑ¯ˆñTôûb¼ƒ”ÄàB”ƒWTšb8Ãsd[m#C£\t}1|‰QJŒ²b´ú•~£™µÂ6íè$ÆémÅ±-Æ+1^ŠILLúbÒ“˜´Ä¤.&U1yˆÉ;¢±ºb:S^NLbºÓ£˜ŞÅô!f1«‹ÙLÌVb6³‘páæb˜}áöÄ4&Ü½pŸÂõ…{ó„˜7Ä<#æ1Šù\ùİ(jrsO,±hˆE],šba‰Å\,±Ø…íóeJ,“bñVZB,Ób™Ë¼XÅ².–U±´Å²%–]±ˆå0Ò¹×b¹+M¬Šb•+]¬bÕ+[¬±:‹•+V±.ı*&‰õC¬’b«\z§Äú±œ'±E,ç^¬Wb½›®Ø4Å¦_±i‰MNl*bSën|msb[Sl‹bs[=*µÏÅ®,v5±†Éö}?ıîÒbWû¾Û«Ø5Ä./¶±‰İLìQêò*vK±óÄŞûD”*Š}LìKbßûØÅŞû‰Ø¯Äş,	q(‰ÃXÖ‘C¾ò"?F”è^6âpOcÆ¢‰c3*ÄwÄq¨VOqÜ†9¢°%sÇ‡8é¿BCâ¤‰SGœêâ”ç”8çÄi#Nqz‹ÓYœuq>FÎyAœâìŠó&ÒÎGâR—Œ8·Å¥	¥mqyDvÙH\ÎâšO\–âæ‰Ë0TÎ®q¹EÄÙN\›âÚ×±¸:âú×CdÅ¤Ä-+n£(f_·}³ÏDéú¡¸ëâ>RÇgq»‹{BÜ®âv›Ç÷™¸w£P¨#î¶¸·Äı.îqŸ‹û+ì„=âZwOúqGM¦Œx˜âQ’xTÅ£&–x¢ GG¼nâõø'Ÿ'Ì*›aÀÉ+¯$¼²ğêÂ«	¯)¼ğúÊ:tÖ›ğµ(%˜¾!ü´ğMá×„ï¿+ü–ğGÂß
ÿ ü»ğ7Â¿ÿ&×{ñN‰·.ŞYñ®‹wM¼†3ï£xßÕ¨#cºŒ%e,'ckÊ˜-c‘ûseì-cOÏÊ¸.ãoÊxKÆ{2>”qWÆ2~ñŒ_ÂFZÜ“ñ·Ô’RÓ¥–’Z[jU©µ¤¶†^’ÚJ&Ú21‘	[&º2‘”‰–LX2™‘‰­Lìe²-“s™ÉäV¦r¿²6ª•“)]&}™ªÊTY&^2µ—©LMdj*S§©3¤•zLêšÔû‘ç3©ŸTOêU©Ï#gà,u_ê7©{2“iMê™ÎÈt>Š¼Ú2İ”iC¦[2İ‹ê)3™~ÈôS¦}iÄ¥¡I#%´4²ÒÈK£JAÆ\KiÜ¤qŒœ[[fš2ÓØ½¨ò8Œ¢q‹°±dldf/3W™¹ÉŒ/³y™­ÈìRf{2û”Ù•Ìdn$s=™«ÈÜIæ2÷…ª42÷”¹»Ì9ÒÔe®#ÍŒÌÙ¿j0Ò<Kó$MKšiîe>&Í»Ì—d>%Í­4—ÒlH³))YhÉÂBz²0”Kš#Y(ÈBZ²2ÿ’…³Ìeş,}YÈÈ‚+‹'YªÊâR²èÉRB›²”“Å¸,^e)-‹}Y<ÈR^–†²Ô—O–n²¦ËâS–×²´¥—¬·d%)ËWY^ÊÊ:Tv+š,ÏeY—¥Uhó–ï²t–Õ¬,#EğƒU;‚Ÿ²Ú
ÁĞÔ#*p kY}GÈR#’‘V‘ÅT—µZˆ,….ÄHÖ›²nüŠ¨„LDS“õ\hA„mòq/K++-WÖ×ÒªI«-i¥¥Õ‰’ä~Xod¤õ¢´î²Q’ºltecU9ï²™•ÍªlæÂ"`ã!›zØl¶ds'›oiçƒ{]Ú5i÷¤=“ö^¶ª²5’­¹l-d+'[¦l]dk+ÛIÙÖek'ÛÙÉÖR¶k²íÈö\¶W²İ“í¶lo#ÌP—Cvb²“íg¨†(q^¶ï²“•“ì¼eç*;OÙ™ÈÎ<D»ıÈzºÉ®¹O'ÙmÈ®‰OÿdÈn5‚¤²Ÿ’ÎP:=éhQ(=•Ñc²WRmK:Ë0îe/-{ºt²·•½ƒìõe¿(ûyÙóvI²W‰jˆ–ì×e(ûÙßËşEöï²?“ı§$å@“ƒŠØràÈÁP–r—_98Ë'‡g98È¡&‡9¬ÈáFm9Êa^Çr¸“Ã·îå(!GI9ªÉQSl9êÊÑ]µÈr<É‘ãd$Yeå¸*'‰î˜ÈñB—Qè=)ÇO9¾Èñ-l*»rÒ—çŸÈÉTNÆQ$o!'Û(¾u•_NQQ5+§U9mF>ä\N§rºŠšñ9}È©ÖÔf19kEşsWÎfrv’³›tÒµ¤[’®)İ¶t»ÒH÷"çq9oËyEÎu97äü%çk9_ÉùP.R?°è…+¹(ËÅD.¦rÑ’‹‘\&å2/—š\¶är)—G¹œEºÈ>bFú‘=rŒ‘¶\™Q¢¾z#aÃ¾$W¹êËÕQ9Ë(½ŒÒÑ®\ÇåÚ”ë¢\ä:•bÛrİ’ëå¯Ò…\÷äz%×[¹>Éµ'7±ŸA²ŠÜTåÆ
3T›–Ü=ä>.wo¹7äŞ”û‚Ü—ä¾.÷C¹_ÉıMîò ÉC&”®YyÈËC!Ìäï_òĞ¤Ñ•‡µ<,äÁ“‡ƒ<åa#y4ä1.cyEz˜'y|F)ë”<e£ZmQòòTR^Aªò˜»†'Kgò¼–ç½<å%./¶¼jòâËËC^‹òÚ’WK^MymËkN^İ‹0í—¼!ïÚ]ysä­$o=y×äíÑ+'yÏş
D„]ÿ{&
ü—äİw[ŞÇa`û¶÷‘¼×åı"ïyÉG\>Òò‘
¡œ|då£*å¨½gËÇB>ò±’|êò±—_>SòiÈgN>Ûò9’Ï—|>ås)Ÿ¾|e#ß»._ù²å«*_ùêÉW[¾LùšÊ×Q¾¡ã¥¤§É—/=Sz™(æ]ˆê¤é¥7
“r~Gú=é¤¿–şVúWùnÉ÷D¾‡ò½ï­ÒÒò}R±”|ßåû¡be««Ø@Åò*æ¨˜¥b9›©ØEÅÎ*v—ïÛ¯®‚ŠU¼¤âå!/Dê¸¯â–Š7UÜQq7‚Ú.ªåªøAišŠ{J‹+-¦´¬Ò2J+*ÍUÚVi{¥Ub¨cf•h«¤¦’•´#Õm§RY•²Tª©R#•š¨Ô<$US¥ç”‰·¡Ò—?‚åaŠõ¤ô‡JÇT:®Ò•Î©tI¥-•î©ô@¥Û*ıRé³J_”á¨ô[5e¬”QWF9$Ğ3•I©LFeÆ!“­+ã­2¹:ÏvTÖRÙjä¸¶BD.{WÙk(È™ÅÒÉ'•¹ù'À´•9Tf+x‘¥óŒØ‡~D@lB¢°S…•*¬U¡¦
wU´T±¡Š®*NT±©Š	UìFÑeCÇªÔ
³²¥ª*%¢‚l)*z.Ué¬JUz¨r\•+ªœTå¶*÷£¦ £ÊeÌUy&ßÊWUöT¹¥Ê/UöÃôZù­,MYiee•US–¥¬¦²leõ•å*k¦¬uhûXÏöñU#Ë÷%‚&òªaªF]5jªa©FG5Q@t¨šÕtUs£š[ekÊ6•]TvNÙWew”íüZñWvWÙ=eO•}PöJµòªÕT­zˆÆ·ª5S­½j"Îâ¤ZÕº©Ö=òäsªQí¼œäU»£:yÕ®«ö]u’ª£GÚ^Uu
ªSV¢êX!Ñ«ÎBuvªsWOu|ÕMFACuµ0¶×í¨n_uÇª;U]¹Šê.Uw¯ºwå¤T÷©¹åèÊÉ)§¨œ’r:Q±~ªœQ±>©^BõtÕ3T/§zÕ[+¯¨z±ì©ŞEõÎ‘xoªşMõÇª?Pı«ê{j «AQrj°Tƒ­ìş)~?Ø‡u½ÁQj¸UÃ½^Ô(©F‰¨Ì:W£¥'Ô8£Æ•Hlªñ@»jìD°R?â;6jì©±¯&ššdÔ$E‚sjRV“JÔ	®«i#=&jºVÓš¾Ô¬§fm5ëFÀ`YÍÖÍ´W³C„{Ü”Šû)5*×VîR¹W5OEiØ˜š'”ë«y]ÍójŞVó†š÷Ô|¨æ“¨8Û~ó…Z¬Õb§¯(Ï_TË²ZVÕ²¦–ZöÔr™ÃSµÜFªáF-O¡j¸¼şÚŒWË§ZéQq¶õ2‡j5R«‰ZÍÔÊJÿ‹°2¸ÚªÕA­“‘ˆXTë¼ZÔº¬ÖÍPGµ¨„Ú”ÕÆR›†Úô"„ÄU›“ÚÆÔ6" !PTÛ²ÚVÔv…ÃÇQ8| vµ=¨í>L‰ïê?É»ªÚ™Q5ÙSÊ»¬Ú%Õn£vmµ›ü(àîvjŸ“¨{]í|µo…éÊ°dyQû›:ÎÔq¥;uÊ„ãq©N	uÒÕi¡NSuÙªÓ(TªÎIuª«S%ôRNïB9ëêlüSjıÜUgGÇê\VçRdÔ%VÉ/Ë(RşˆòÌã(É|P×då¼·0¾{ÕÃ4ïU‹²¬Õ(2:¢avÔS·–ºµÕÍQ÷³ºßÕİSxdgåBáQQzX=tÕc¢sõ¸©Ç"µºõ4Õ³ Å¨EZSÏz:êÙUÏ~ÔvÕs«§0Qù¼ªç="\^êé©WZ½ÑºªW]½êµR/_yYåUB×Å)¯E£Ê›*o¦¼•òÊ»*ï ¼Kí(¿ÿk¡ı-gi·EyŒtGèö¯QLñÈÊqGå•5dµ¼3B{²¹çÈP¹;ÛOx©àæÏAğ|¯XğJ¯{àY_	ülà·¿¼óˆÇï7bâ:â;Ä5h5$NH4‘Ê"5†Ş…î#½‚QG¶ƒì¹rm˜1˜i˜cä(TQ8 ØıæPé ’Eeê–®îÎ
=½ô7·øÍ-aè`xÀ8‡I“6&]LF˜1}`maµ-¶MlWØ98LqxàxÆi†sõ›[ÆYÃyË—;®%¼xàÍàÿøÿøıúo‡ñ.µµ,u&ûL˜,3yerG½I£Êì”Å3Ë}Vº¬X/°c³Îæ“Í›66Ç´'´´íon…ö‘-›­~D[¿Ù~°Sø6³;¥óä Â^Œı8IŠßÜ*'î7·Æi3Ò­wœ/Cs~äÜçrÈå˜[‹»7÷tî·<¶xJò’æyË‹ÅËŠ×"¯7^¼¾x»óŞámÆÛ‹7}ƒş™ïsÄò¤EÌZJhK‘,ŠdR¤:"¥‹Tú›[¸K_d¿Ïşãûõ‹ÈúÂs-‘¿Š‚'ŠQ´Dq$ŠQJ‹âY”¦¢´¥¹(mDy9Ò=Q‹‹úV´´È…PèÖô›Û]Ù­èÄ  c1xˆñî›ÛüæÚßÜ–˜Ö¾¹m1m‹eY,gßÜX.ÅÆùævÅÙç¸\Å¥/®–¸W„—şà›ëH3.óC™ËbZ/²lÈÊ)r mY_~s{ÒÚ}sûßÜì§åè,İŠt¯r~—‹Y÷¾¹C¹¼DÅ;OòT‘çË7w$ŸÅoîø›;	iß„óÍªd]%g*ı£§ûCêK7TÚQéµJoTz¬ÒeèÊ()#«Œ‡ÊÄ”ñR™‚ÊØ*SVÆ]ef*;S™«ÊÜ"¶·ò¿¹Ò7w‰¾}•«Føá@åß\7Cö°ˆù]¨oÆª¨GÆ-Œ£ïd[•o³Î7w®Z]ÕÚªÎQuQ‡8¯ºEÕ-«n=+“oîâ›»TãÙ7wz†ó‚ZÚ¡<¢»µ)ªÍ^mj‹¤ˆå7w­n¶z–|×ÿe.S»ïµÕ÷ğÛ{üí•şöúùÿñ?ÿûÿë”oş=ıÛíïöß»¿mÿŞş{ë›çü‘ü#õç'ı¯0DúÛà·éoï?ıÿÿ{ü¶ømşÛæ'‡¿ÿwÒúùÑ*£èdşÈş±øëø¿…ÇÍ¿^ÿ¿áëÂË¿„µÔ?Š¾¾…?V¬ÿØü±ıGöÒ?ÌhÿÈÿ#÷Çî…ŸŸõoÑWøŸÚßşå¿ı_ÿò·ÿ,üçŸşß¢ˆêU¾ä¿œÿêıüÈ÷_Âúgé¯#ÿ÷ßşåoŸÓ¿^ÿkøúğ×ëÿñ·ùÛWüóı×‘ğ,|e?š_•¿…§ôkúáüy$ª¤~L?Ë»ÏÙçşóøyûÒ¾b_™¯ÜÇí«úÑûš}ø??ûû¯½øGì#ù¡}j¿ç?cŸÉÔGâSÿ4>Íôgö³ğ‘ù0>rÙÏÖ§õYùlşù§£dígçÓù~ö?Çùó£øQø½ø¹û\|n>İÏÓçåóşé}>¿_é¯ÔGù£úQù°>j¿—¿J_æGı«öÕü²¾Z_¯ş—ó5ş~´>ìÎGûkõ{õk÷µø:|m¾Üß¿··¯>º_·¯Ë×ıËûzşüJ¾ÿbú|ô?ÆÃÏÄï…Ïøgêcò1úLf>ó³ÏÜgñcşá~,?ŸíÏÆgõÓşıüÙıüógôïÿşÿ…ßÍàsò±şX}l?6¿ÿùÓú÷ÿõÿ„[~n?çŸçÏëçãÓÿ|}%¿Œ/ıcÿqü8|\>N¿W¾Ê_ùóWıËşj|µ¿º_ƒ¯Ş×äkôñø¸¼>_ëßk_û¯å×ñkû5ÿ½ù{ç÷ÖïÖ×ùÃûx]¿_ş×ëÇßù³¶zÓ‚[2¸¥‚›Ü*Á­ÜFÁm¾Yß–ÁmÜ¶ÁmÜÎÁ=ÜãÁ=ÜÓÁ½Ü[Á½ÜGÁİîóà~
îçà~îïà‘VğhnğØ{ğLÏRğlÏvğìGË€ağtƒç6xîƒç1xƒ§<ıàù–Z´B0‚W.x™Á«¼êÁË
^àe¯MğÚ^?xƒ×9x]ÃµÄë¼üÀ‹xFàe/xfàU¯®4¼fàõox‹À[Ş&ğöw¼[àİïx^àù|-ğ“?ÏÓûQìK~&Zª˜_üR´r©†Ë¿øv´„q¿ø‡À¿ş=ğŸÿ
|/ğıàŞñà
Şzğ6‚w&xç‚·¼óÁ»¼+Á»¼ûÁ{„õCğ~o/ZÅ‹#–D,X1±b%Äªˆ5k"ÖF¬ƒX7\1Å†ˆÍÛ ¶Cì€Ø±s¸¤ŠÇÃeT<-©²ˆ›ˆ/#^A¼Šxqñâ}Ä‡ˆŸ">C|ø:Z‚í{B‹CK@KBÓ¡e¡ ¡ıü7ü£¸
­­®Ø4ZZšÅ¡- m¡í İ¡= yHÄH ‘D"„‰D‰U$jáb/ÑA¢‹D‰!c$¦H¸HÌ‘Ø"±Gâ‚Ä	É’:’i$$sHšHæ‘, q@²aì6’íÈfï"é 9Br†¤‹ä*²·rxü%o¡’Ò¢¥f©*Ru¤H5‘ê…‹ÏÔ)zzzzz3Z‘  ¯¡ß ? ¿Â5ªşFúçûÓØ*ÒÒE¤ËHW®"]GºRñé!Ò¤7H>"}Bú
#Cƒ‘‚QQ¼†£	cÃ…±‡q€q†ñB&…L™:2dšÈ´é ÓC¦Ì ™)2d®ÈÜ¹#CVC6…lYÙ*²õĞZË6"6¦-«ÇÈNÂÅuv†ìÙ²d¯ÈŞ} ûDÖ‹ÖİIätääLäŠÈ•‘«!g…ëñ\¹rä\äæÈ-[!·CnbÀ¹KH»…köÌ,ÌÜÏóôãzÃ4a–aVa6a¶av`:0{0'0g0]˜s˜K˜+˜k˜;˜˜g˜/˜Lyyùò6òmäÈŸ ¿B~^!äïÈû($PH£` Pˆ®j(Ø(tQ 0FaÂ…mx-Q8¡pFáŠÂ…
Š1S(Q¬ XE±bÅ.ŠŠ}§(ÎP<£¤¡TE©…Ò ¥J.J”¶(íQ:¡tF9rå4Êy”‹(;(Q£<Ey†òå%Êk”7(ŸPşù†ı#¸ŠJ•*)TT2¨äQ)£RGÅŠ.|z¨P™¢2¯€*TÖ¨lPÙ£rEåÊ•'*oTuT³¨æQ- ÚGuê
Õ-ªT_¨úáÕS-†ZµjjÔLÔò¨P«£¶@m‰Úµ3jÔ¨½PóP£F=ƒzõ<ê%ÔË¨7Qo¡ŞFİA½ú õ1êSÔß°â°´ğ2Í2``a•aU`UaÕ`5``a­a`½Ñˆ£‘@CG#ƒFe4~.b~$WÑ¨¢Ñ@£‰F!s4.h¼ÑL¡i ™E3µÜÑ´Ñ\¡¹CóˆæÍ+šw4Ÿh¾ĞôĞ|ÃÁÃNÀNÂÖag`ç`çaa×`;°û°°‡°G°§°ç°×°·°÷°°Ï‘BË@+¡ØõĞTı«:hõ~Z‡.ZóH?\E â>ôÎZ×Pt
½ÚÙPíiÑ.¡m†øHÛF»ƒvíÚ´¿·ÙÑŞ£½D{ıó´üxH¢}AûŠN4:YtLtÊèÔÑi¡ÓF§‹N!:#tÆèLĞ™£³@g…Î-:GtNèxèøè¼ÑM¡«£k¢›G·‚î0¼ÚîÎĞuÑ]¢»şşÕtèĞ=£{E÷îİ'º/t}8q8	8&œœNNN§g §gÇ…3¯ä½ttIo¢WD¯^=½1zSôfèÍÑ[£·CïŒŞ½zz>zoôè'Ñ7ĞÏ 7øy‚~ä/ú&úyôÛè;è÷Ğï£?E†şıúwôßd001ÈcPÄ „AƒÆÌ0p1¸ağÂàaÃ8†i3f141ÌcXÄ°„aCÃ.†}1œb8ÇpîNŒ’¾0Ò020ÊcTÀ¨bé¡šnıô01šctÄ8…±îiŒÍÉW1®I;qãÆ#ŒÇ»o1Şa|ÄøŒñã7&	LtL*Ñfˆıs?ÄÁ¤‡É Ü™,1ùóÒÀêdòB“3&Lî˜<1yašÃÔÄ´€iÓ:¦L{Ñ¶ÊÓ5¦L·˜î0=bz÷Z¦/Ìâ˜%0Kbf`–Å,Y³.ff=ÌF˜1›a¶Ál‹Ù³f'Ì®pcp“pM¸y¸¸¸U¸ÜÜ1Ü\îîîó$æ)ÌÓ˜˜›˜—0¯cî`>Â|‚ùó5æWÌ˜?1÷±ÈaQÂ¢Š………E‹},†XÌ°Xa±Çâ€Å‹37,X<±øó÷õ= eKKË"–-,;Xö±œ`¹ÀrƒåË–V1¬âXX5°:aecÕÆª‹Õ «	Vk¬nXÇX]ëJ¸Kµv°v±Ş`½Åú…µ‡µõ›86Iltl²Ø°)aSÆ¦ŠM›96+lÖØl±ÙcsÀæŠÍ›¶1lãØæ±-bk}ÿ¿»íc»Åví	[;»F¸-¶ëa7ÂnŠİ»%v+ì6Øí°;`wÆî‚½}ûöCìGØÏ°w±Ÿÿ<A?êªû5ö;ìOØŸ±`ÿÄş…½‡C‡,E*8tqèáĞÇaˆÃ(Ú“[àpÀá„Ã9ÚŸËà˜Ã±‚cÇ÷8p<áÇIÇ)S§%N#œÆ8MÃı¼Ó§Nwœ<œŞÑÆ^ggç:Îeœmœ³8÷pîã<ÀyˆóçE¸ÿwŞá|Çù³ó—8.	\\²¸Tp©âRÃ¥‹K—.}\f¸,qÙárú¹}˜Ëê+áZÀµ‰«k×ñÏóô}×9®7\ï¸ÅpKã–Ç­‚[·:nMÜz¸õqÛàæâvÀíŒ›{÷îÜË¸[¸7qoãŞÁİÁİÅ}ñıç{ßáÃ#‡G!c<&x¸x,ğXá±Çãˆ‡‡‡ÇOOÏ<ËxVğ¬áYÇ³ƒ§ƒçÏ%^^:^i¼¼
xÙxuğêFÛ¤C¼FxñrñZâµÂkƒ×.Ü>}]ğzâåáõ†ƒ—„—‚—†—‡W‚Wƒ×„×†×‡7‚7†7ûy‚~¤U½-¼¼¼3¼¼<??¿ß‚ßß…?„?ƒ¿„?†¿‚¿…¿ƒ¿‡Ã;‰woïŞ9¼KßÏÑÛÁ{€÷oï9Ş¼Wx¯ñŞà½Åû€÷ïŞ¼_Œ¥Ë2–g¬ÊX±cMÆZŒõ3¶fìÈØ™ñãiÆÆ3Œ›Œ—¯1n1ŞwãÆgŒï÷SÓ¨éÔŒp+Y+P+S«R«SkRëRs¨©©M¨¹ÔvÔöÔ®?OĞÇ"´'5Ÿ	‰$&²Lä˜(0Qd¢nP'l&ÚLt£ıê!“p¿:qfâÆÄ‹	IÉ4“Y&M&¿Ÿ&-&LÚLv˜t˜2¹`rîp'÷L™¼DŞw¦âL¥˜Ò™Ê0å0ÕgjÅÔ©SO¦^L½©Ç¨Ç©'¨'©ç¨ç©W¨W©7Âır½C½K½G½O}Bİ¥¾ ¾¦~£~§şb:ÆtœiéÓE¦ëLw˜î2íDVı€é!Ó³Ğ§OO~ OD¤W‘e¿ıKÊ3Òá½Q§Ñ¡Ñ§1¤1¢1£1§±¤q§á1£1“`&ÉŒÎŒÉL‰™23fjÌXÌ´˜é03ffÍÌ–™#3f<fóÌ˜-2[
 Y—Ù#³f¯Ì>˜õ™}3—d®À\9‹¹sæÌ˜{1çÑLĞLÒÌÒÌÑÌÓ,Ñ¬Ò´h6iÎi.h®ii^hŞi¾˜3¯1Ÿ`>É¼Î|–ùóæ+ÌW™¯1ßd¾Íü„ù)ó+æ×Ìï™?1şq~&Uó7æ_Ì{Ì¿YĞXH²f!ËB‰…ulZ,8,ŒY˜²°faÃÂ–Å8‹),–Y¬²Xc±Îb‡Å>‹,®X<°xÇ"Å‹w–2,eY*±Ta©ÆR“%‡¥>KC–Æ,MXš³´`iËÒ¥3KKo–s,Y.³\a¹ÊråËm–;,wYî…3—ò€åËk–,ŸX~²ì±’b%ÍJ‘•+eVš¬Ø¬´ÃM¥ÇÊ€•+SV\VV¬YñXñYùùŞÿ#ªÊj’Õ«YVs¬šál§ZbµÌjÕ«cV'¬ÎY]°ºfuÇê™Õ+«OVß¬ÅYK²fD„y–µbH×jol‡˜kÍa­"!Õ9­¾Úµk/Ö<Ö¬§YÏ°^b½Âzõë}Ög¬/Y°îÑŠÓJÑÊÒªĞªÓjÒêÒêÓÑšÒÚĞ:Ğ:Ò:Ñz³¡±a°‘a£ÀF‰ÍPYnìØ8°qbãÌÆŸw4ÂJ²iD#¬<›e6+ÑPËb³ûó<ı`{ÂùÖ€Í!›£pĞÕœ²9gsÁæ’ÍU4 Û…ÃÅæ‘Í;›WÚ1Ú)Ú:í,m“v™v!š5iÛ´Û´;´û´l?i/h/i¯hûáØÌ>Ó~Ó¾Ò¾Ñ¾Ó~²£ıb+Ç–ÉV­[•hºÖb«ÍV‡­^4i³µfkÜÎl=£Á[ŒmíÛÛ&Û¶‹l×Øn³ıŸ^²½
QÈpD—b'ÍN*;6;vúìÙ9²ã³›b7Í®Á®Énİ:»§ï{Iát¯»bwÃî‰İ»O:q:I::E:e:U:&›N›ÎÎÎÎ9:/:{÷p8ØÓØK°—bÏ`/Ã^…½{6{öìMÙÛ±w`ïÂŞ5œ'ökìÛìwÙï±ï²?cÅşšı-û;ö÷ìÙ²ÿgƒYLê48hrĞæ`ÉÁæûûÅàÌÁƒO^j&9Ô948´8lrØâ°Ãá€Ã‡ç.8ÜpxàğÉQŠ#ƒ#“£Ÿ³’MUJ•9ªrdqdsÔá¨ËÑ„#—£GKV­9Úq´çèÌÑ…£'G/<u3g9.q\æ¸Ê±Åqƒã=ÇGÏß8~r¢q’â$ÃI“'mNFœŒ9™p2åÄådÎÉŠ“'{Nœ<8yròæ4ÆiÓ
§V8ˆ8p:åÔåtÉéšÓ-§WNoœ>8}q–âÌà¬Ì™ÅY#ÜÎ&œm8Ûr¶çìÈÙ™3ŸnŒnœn‚n‘n™nn‹n—nŸî˜®{İŸs UUº'ºWºOº>İ7ç©hœá<ËyóçeÎ+œ×9·8os>æ|Æù‚óı©ñ"ÎE‚‹4.r\”¹hsÑåÂá¢ÇÅ„‹%[.v\ì¹8qáqásñæRç2Ã¥Á¥Ée‰Ë—u.-.;\vÃaôrÆå*Û÷\B³zé…xõ*Á•ÎU:‰WE®J‘^ÛàªÍ•ÃÕ˜+—«%W+®¶\¹ºpõäêÍµÎu–ë<×®-®›\w¹îs=ãzÉõ†ë×g®½Ÿçé_ø>77)nÜ´¹éqÓçfÈÍˆ—›97+nÖÜl¸ÙqsãæÎÇÏm‚Û$·nÍp´¾mq;àvÄí”Û·n÷Ü^¹½sûæNã.É]Š»Zæ®À]‰»
w6w=î¦ÜÍ¸s¹[q·áîÈİ™»wîÑĞ>Ë}û<÷Eî+Ü×¸·¹o‡Ãüı’ûu8ÏßŸ¸¿pãşÎı“‡T8ç?¤yÈğPã¡ÎCƒ‡&z<ìx8ğpâQãQç1Ïc‘Çe+<Öxü9-ıQVo8vxìó8ãqÎã’Ç<Â;	Ni²<x*òTæ©ÊS‹§)O3\îg<mxÚò´çéÎÓ‹§7Ï{<xğ<ïE8ïx~ñ¢…·&\Ê¼Tx©E7(4xiò²ïT¸œx¹òrçÅãåÍk’Wƒ×,¯…è&‹×6¯]^{¼öyò:áuÎë–×¯'^/Ñ-1Ş’¼9¼£¼]ÂûnÏğ¦‡ÿÃØ›ôº®,kbÅØ“²¡°Ô¬%©ŒòÄ3ì‡·ŞàÜwÏ±=ráİ÷ …¨¢$ª§z’êû†êE‘pkàÙóõO°–2#b²÷_0H­}Î´ú@“™É`d¤˜Œøx‰óK‚_D~IóK†_òüRä—¿”ù¥Â/5'<âÒá—ŸïQ¿¸Uù¥Ç/}~™ğË_®übs3ÇÍ27ÜÔ¹9ææ’›+nÜÜqóêW\~ğkŒ_Sü*ò«Ä¯9~•ù5Ï¯s~]òëŠ_/üzåW›[qn‰Ü’¸•åV[
·êÜjr«Å­·¦ÜšqkÍ­·vÜ:rëÊ-‹[Knç¹]äv‰Ûen×¸İàv“Ûn«Üîr»Çí	·Ü¶¹}ã·¿¥Ü¨¿eù-Ço~+ñ[™ß*ü¦ğ[“ß4~ÓùmÈo#~óÛ„ßVü¶æ·+r;ò›Åï	~Oñ»ÈïÒO=}ÑoÜ›ü®òû„ß§ü¾á÷¿_øİâwÁ3Iº‘&9j è ôAXƒ°áÂëç·í_¶— ’HÅı„y"*DdˆœoúFYAä ‘#Dn¸¡+IˆVÜ™
¢-ˆœX–è¢D÷=AL‚XbeˆU V…˜±:ÄÚ@Ì„xìuj<ñ¬£B\‡øâKˆ¯!¾¸ñ3$"H¼bc !A"‰$J¨A¢é†Í¨Ğjæ‹Q51Ä’$LA2É$›lAR‡d’kHn ¹…ä’&$ïÜ¸›8¤2ÊBª)RH5!Õ†ÔÙ	ÉS Š ¦A,ƒX±	bÇùÊ»8 qâÄ=ˆg/@:
é¤–!‡tÒeHW!­@º	é6¤UHëîCzé9¤W^CÚ€ô	ÒgH_!mAúR¤4H9dÊ µ@š€´i’ÒŞ	ÊD “‚L2YÈT!£@¦2CÈì ³‡Œ	™Ÿïî¿8U!sƒl²qÈ& ›‚l²9È– [†l²ÈªíAv Ù1d']@Ö€ì²Èš½BÖ†\r9È W‚œ¹:ä:S!§AN‡ÜrcÈM!7‡Ür+È­!g8aL¹ä,…¯¯¨Ë	“ Ë @.ƒ\¹rä6È*ÈÈ:È=‡ @<y	òäÈ-'.J¾C>ù¤ •|òy'F*_vã¥ï@¾ëN ?†üò{'|*‚üò&ä¯ÿ}†{­*Aş
(Äh«‚…"šPhAaàF^MÜ(\ (@1Å8“PLA1Å‹P,A±ÅœUAqÅ	çP\BqEŠ[(	PŠ@)¥$”RnWJ(¡T‚RJ(5¡Ôvb»J:”ºPêCi¥	”¦PšAi¥”ÖP2 ´…ÒJ(]¡t‡²å4”³P.@¹å:”›î7×U(kP@yå”çP^CyeÊW([P¾C%ùS#¯e$¨” ¢@¥*]¨ 2†Ê*3¨¬¡r„Ê*6TnP¹CU€jª)¨ŠP•¡Z€j	ª5¨ªPÕ ÚubÔªT/î‹&P" $@)€R¥J”(mPz @Ùƒb‚r‡š µÔâPKB-5jy¨Õ Ö†š5j]¨ 6†Új+¨­¡f@íuêQ¨Ç ‡z
ê"Ô%¨g¡‡z	êe¨W¡®@½õÔÛPï@]…zêC¨ >…úê[¨ï ~€úê—ŸÚùbN­›P· !AC€F1hÄ¡‘€FhT¡Ñ€†4†ĞXBc54hì¡fšYhæ YuCöÚĞì@Sƒ¦ÍŞ7Ï‡ç4gĞ\BÓ†VZ	h¥¡•VZ2´*Ğj@K‡ÖZh- µ–­#´.Ğ² -@;	m	Ú2´KĞ®@»
í´ĞnB»í´'Ğ^B{í5´÷Ğ>@ûí´mhß …N:IèHĞÉAG†N:%è(Ğ©A§:ttèô~ªæµztÆĞ™Agt¶ĞÙAg:6¨¨qP³ æ@­€ZUµ	ªjÔ1¨3PW ®A5@=€z-Z´"h%Ğª 5@SAÓ@ÓA€6mÚ´hh{ĞÎ ]AO€=ztô"è'RïŞ}	ú
ôèĞOĞA7İtEèfİ˜ÉtÛĞU¡«C·İ¾E9îº+'œ²{ƒ½<ô*Ğ«C¯½ôTèiĞ@o½ô–ĞÛ@Ï€Şz»/=}1¶Bï
=zwè§ /B?ı,ôeè— _†~ú
ôĞoA_ƒ¾ı.ôGĞŸAıè¡‚¾	ı+ôï0HÂ ƒš0hÃ@…CŒa0…Áö08Àà
wÆa˜€a†30ÌÃ°Ã
ë0lÁ°Ã1§0œÃpÃ¯0´`xƒQFYaT†QFU)0ª;Q¥£6Œº0šÀh
£9ŒÖ02`´‡ÑFWGa‡q
Æ"Œ%ça\±ã&Œ;0şùgñ‹³Æ:Œ{0Âxã9Œ—06`|†±“(Lb0IÁ$	&Y˜aR…I&*Lº0Ãd“%LÖ0ÙÀd“LÎ0¹ÀÄ„i¦	˜¦`*Â4Ó*L˜6`Ú†i¦Lu˜a:†é¦K˜`z†éfÌ¢0KÀ,	3	f˜`V„Yf*Ìt˜õa¶€Ùf˜]`v…™ó$ÌÓ0ÏÂ\vãn[0ŸÀ|ê„ŞÎ07`¾‡ùæ6Ìo°HÁ¢}XŒa1…ÅsX,a±‚Å›ŸzúbÕX°ØÃÂ†Å–,#°LÁR†e–:,°Ár
Ë,—°\ÃrË,O°ŠÃ*+VXa¥Àª+V:¬†°ZÀj«¬¶°:Àê+VWXY°²a…uÖ"¬%XW`­Àºë¬›°nÁzë¬·°ŞÃú ë3¬/°6amÁú›l¢°‰Á&›$lÒ°ÉÀ&6yØ4`Ó‚›!lF°™Àf6;Øìas„Í6&I0$0ª`Á(Q£ñÍóş¥š/*£FŒ1S0–`¬ÀØ€±ã††Û$lS°ÍÃ¶[¶:lG°İÀöÛlMØ^akÁö»8ìR°a—†]vyØ•aW]v5Ø5`×„]v#ØMa7‡İö1Ø§`/Â>û<ì‹°/Á¾{öØ«°ïÃ~ûì—°7`oÁŞ†ı"pH¹1Ù98áP†C84(íCC8Œá0ƒÃg8˜p¸ÂáÇ£pŒÁ1Ç$E8fà˜…cE8*p¬ÁñçÂÚY+U8à8†ã;8İ8ğœ¢¯ÿò§,œrp’áTƒSÎœzpÂi§œæpZÂÉ€ÓNG8]àdÂY„sÎE8Wá\ƒsÎ8ëpÃyçœ—p^Ãyç#œïp‰Ã%—\d¸œğóK.\ºpéÁe—9\VpYÃe—3\n`FÀL‚™S3fÌ*˜u0›`vÀTÁ²Àº‚eƒuëvì¸ÕnçÁ.]»vìØ]°û`À‚=ûçÛş/V°°m°o`ßá&À-·Üòp+À­7n5¸µáÖ…[n#¸Íà¶„ÛnÜîpàƒ»wîy¸·à®Á½÷!Ü'p_À}÷5Ü·p7ánÃı÷;
)D$r(È(äQ(¢PF¡‚‚‚B…
#(¬PØ °Gá€Â	…
&

7Œ‰cDÇÈ ##Œ,1b`ä€ÑFEŒÊmcTÅè£cŒÎ1zÆèc"ÆrËc¬„±
Æª«a¬1cKŒ­~êéëC<±Æ¶Ûaìˆ±ÆLŒİ0Ãxã
Æk¯c¼qã:Æc|ñÆ71na"‰&$Ld1QÂD:&º˜cb‚‰)&–˜Xcâ„É&˜Œc2É$&3˜”1©`²É&ULö0¹Âä“{L^1•ÁTS	LI˜’1•ÇTS5Lµ0ÕÁ”Š©.¦˜šcj©¦Lm1uÀÔSgL]0ecê¢€bÅ<Še+(6Pl¢ØA±‹bÅ	Š[÷(š˜a:é¦şÉşâjÅtÓ2¦ó˜.bZÁtÓ:¦û˜cz‰é=¦¯˜¶1}GI@)‚R¥8J)”Ò(I(É(Qª ¤ TG©‰R%¥.J}”(Qš 4Ei†Ò
¥J[”ö(P2Q²Pº¡tÇL3ÌT1ÓÄŒ™f˜™`f†™%fÖ˜9bæ„37Ì
N"H6…Y³f3˜«`¶…ÙfUÌj˜íb¶‡Ùf—˜]avƒÙfO˜51kaÖÆ\s1ÌÅ1—Â\s2æŠ˜«bNÁ\sÌêé‹M#g;i'²€rå8ÊI”s(—P® \E¹råÊm”;(P£<Eyòå;æ7e%ù<æ˜/b¾‚ùæÊ˜Ÿ`~æd³äÏ˜¿`ŞÂBI,HXÈc¡ˆ…*(/±°ÆÂ‹	,¦œÔ—¢„ÅóX,b±„E‹u,6°ØÄb‹:XbqŒÅ7XÜaqÅ#ÏX´°hc)‚¥–JX*c©Š%K#,M°4ÅÒKk,N‚MÉÄ’¥–,G±œÂ²ˆå,–óX.ÿÔÓŸF¹å1–gX^cyƒeË[,Ÿ°²Å²‰å+–m¬D±ÃJ
+Y¬ÈX)`¥+=¬ô±2ÀÊ+c¬Ì°2ÇÊ+;¬ì±rÄÊ«I¬æİ¹ «Š›÷ÓvŸV°:Àê«k¬^°jbõŠU«wT"¨$PI¢’F%ƒJ•<*eTê¨tPÑPÑQé¡2DeŒÊ•#*'TÎ¨X¨Ø¨ÜP¹c-ŠµÖ$¬e°&£²ÄZkM¬iXÓ±6ÂÚkK¬í°vÄÚÏPÿ/vV¬İ±.`=õÖsX—±^Àzë%¬×±ŞÀzë*Ö»Xc}Šõ#ÖÏX7±naİÆ†€(6’Ø(a£Œ5lÔ±ÑÀFúK)6ØXcã€#6NØ0±™Ãf›
6ëØlb³Í.6GØcs‚Í%6lî°¹Çæ›'l±yÁæ[¶"ØŠa+­"¶JØªb«†­&¶ÚØÒ±ÕÇÖ[clŠØ2±eaËÆvÛ	l§°Á¶Œí<¶Ø.c»‚í¶UlkØîb»ÿSA_Äí	¶WØ^cÛÀöÛl›Ø¶°mcû†;1ì$°“ÃN;ìt±3ÀÎ;kìØÙbçŠjìejÕ,ªTË¨VQ­¡ZGµªŠªjÕ1ªT—¨®P5P= zCõZ5µ4jj2jUÔÔ¨5Qk£¦¡ÖC­Ú µjsÔ–¨¨íP;£f¢vG=‚zõ$êiÔ%Ô3¨Po ^B½Œzõ:êMÔ[¨·QWQ×PŸ ¾D}…úõ-ê?ßğ±±¢n¢na7‚İvØ±›Æ®„İveì±[Å®‚İ:v›ØmcWÃ®İv§Ø5°{Áî{öêØkb¯…½6ö4ìéØ›`o†={6önØ°/½L©_Ä~	ûeì·°¯bŒı5:[5œœºş	$qÃŒƒ<ª8PpPÇAM´q ã`„ƒ1&8ØàÀÀÁ\p`âà†C‡Q&p˜Äa‡y–pXÃa‡-¶qØÅa‡}Îq¸Âá‡;phşTĞQÆğŠC‡6Ò8Êá¨€£
ª8já¨#G:†8Úàh£8:áÈÂñkeÇCOq<ÇñÇw8>âø‚cÇwœÄq’ÀI
'yœİ„Â
Nª8Qq¢ã¤‡“NF8™ãd“5NNÆáÄÆÉ§Nc8Mà4…Ó4Ns8•qZÆ©‚ÓNë8mã´ƒÓ.N‡8ãô€Ó#Nİ„Å©‰Ó+No8½ãLÀYgIœep–ÇYg%œ)8«á¬3ÍÍnœâl…³5Î~Ê-áìˆ³Î®ndçIœ‹8Oã\ÂyçeœWp^Åyçœk8_à|ƒó=ÎNªäÜÂù	\$q‘ÁEm\¨¸Ğp!áb‰‹.N¸°pùúK—%7Å²‚Ë.s¸TqÙÇå—G\qyÁ¥…Ë®"¸Šã*«®r¸Êãª‚«*®Ú¸Rq¥áæŒ›nlÜÜĞH ‘B#F<E4ÊhTĞ¨¡QG£…F}4hŒĞ˜ 1Ec‰ÆG4®è|øé÷¿@__×Ù¦p›Ám·%Ü–qÛÀm·mÜp{Á­‰»î’¸Kã.ƒ»<îÜUq×À]w]ÜMp·ÀİÚM'İâî€»îÎ¸»áîûîc¸O;i¦ûî[xâaŠ‡Vx¸ã1ŠÇ8E<æğXÀcU<ªxÔğ¨ãq„Ç9x\ãq‹ÇOx¼àñŠ'O	<%ñ”Â“ˆ'	O2Šxªà©Š§xàiˆ§1xÚàÑÂ“§-öx²ñÁsÏ	<'ñœÁsÏ<+xÖñÜÅóÏeµ/V<÷ñ<ÂóÏwx>âù„ç3/x¶ñ"à%‚—^rx‘ñ’ÇK/-¼´ñÒÁ‹Š—.^zx™âe—%^ÖxÙàÅÀË/&^®hFÑŒ¡™F³€f	Í*šM4[hvĞTÑì¡9FsŠæÍš;4hšh^ñ*à5‚×^ãxMá5×"^+xmãu‚×^7x5ğzÂë¯&^—h	hEÑJ¡%¢•G«€–‚V­:Z*Z:Z]´úhÑ£5AkÖ­Z{´hÑ:¡e¢- ù©§×rÚ1´ShçÑ.¾„]G»ƒ¶Šöo¼éxëâm€·!ŞFx;àíˆ·3ŞnxğÁ{ïŞ3xÏá½€÷ŞËx¯â]Ç{ï¼Oğ¾Â»÷Şx¿âıFB”„8	"	Yê$´H˜0%aMÂŒ„	Kv$ìI8‘p§HŒ"qŠd)Ò¢H‡"=ŠŒœÔçÈ‚"Šì)r¡h‚¢"E3ÍRT¦h¢UŠÖ(Z§hƒ¢-ŠêíR´GÑ