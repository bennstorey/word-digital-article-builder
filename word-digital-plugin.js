/**
 * Word → Digital Article Builder — WoodWing Studio plug-in
 *
 * Content Station SDK plug-in that converts Top Gear AN+ Word documents
 * (.docx) into digital articles (.digital). Two entry points:
 *
 *  - A button in the Dossier toolbar (the happy path): parse a Word doc and
 *    create the digital article directly inside the current Dossier via the
 *    workflow API (upload through the Transfer Server, CreateObjects with a
 *    'Contained' relation, C_HEADLINE set from the feed headline).
 *  - A Custom App in the Apps menu: same converter with a .digital file
 *    download, for when no Dossier context is wanted.
 *
 * GENERATED FILE — do not edit directly. The conversion engine is extracted
 * from index.html by build-plugin.js; edit there and rebuild.
 *
 * Registration (Studio Server Management Console):
 *   Integrations → Studio → Plug-ins → Studio → Add new → URL of this file.
 */
(function () {
  'use strict';

  if (typeof ContentStationSdk === 'undefined') {
    console.error('[word-digital] ContentStationSdk not available — plug-in not loaded in a Studio context.');
    return;
  }

  // ─── Conversion engine (generated from index.html) ────────────────────────
  const TEMPLATES = {
  countdown: {"version":"2.4","data":{"content":[{"content":{},"identifier":"header-image","styles":{},"data":{"an-image-role":"photo"},"id":"doc-1jj1hvk1o0"},{"content":{},"id":"doc-1jgpus9k50","identifier":"title","styles":{"text-align":"_align-middle"}},{"content":{"title":[]},"identifier":"subtitle","styles":{"text-align":"_align-middle"},"id":"doc-1i1hesi1t2"},{"content":{"name":[]},"identifier":"author","styles":{"text-align":"_align-middle"},"id":"doc-1i1hesi1t3"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih890"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"50"},{"insert":" Xxxx"}]},"id":"doc-1jgpumpn20","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih892"},{"content":{"apple-news-follow_tagline":[{"insert":"For more content follow this channel"}],"link":[{"insert":"FOLLOW","attributes":{"href":"https://apple.news/TRtxUc-x1TtSqnoPE8MyWFA?subscribe=1"}}],"image":{"id":"72516","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"apple-news-follow","styles":{},"data":{"an-dark-mode-image":{"id":"72513"}},"id":"doc-1jhtp4fs90"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih893"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih894"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih895"},{"content":{"text":[{"insert":"49","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpumtrb0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih897"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih898"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih899"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8910"},{"content":{"text":[{"insert":"48","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpun0sn0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8912"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8913"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8914"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8915"},{"content":{"text":[{"insert":"47","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpun3hk0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8917"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8918"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8919"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8920"},{"content":{"text":[{"insert":"46","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpun6ch0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8922"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8923"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8924"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8925"},{"content":{"text":[{"insert":"45","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpun9dp0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8927"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8928"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8929"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8930"},{"content":{"text":[{"insert":"44","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpunc360","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8932"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8933"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8934"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8935"},{"content":{"text":[{"insert":"43","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpunhei0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8937"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8938"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8939"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8940"},{"content":{"text":[{"insert":"42","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpunk2d0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8942"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8943"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8944"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8945"},{"content":{"text":[{"insert":"41","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpunmpp0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8947"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8948"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8949"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8950"},{"content":{"text":[{"insert":"40","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpunphh0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8952"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8953"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8954"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8955"},{"content":{"text":[{"insert":"39","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpunth40","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8957"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8958"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8959"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8960"},{"content":{"text":[{"insert":"38","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuo0cq0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8962"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8963"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8964"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8965"},{"content":{"text":[{"insert":"37","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuo4tr0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8967"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8968"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8969"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8970"},{"content":{"text":[{"insert":"36","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuo7q10","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8972"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8973"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8974"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8975"},{"content":{"text":[{"insert":"35","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuoaii0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8977"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8978"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8979"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8980"},{"content":{"text":[{"insert":"34","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuodkc0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8982"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8983"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8984"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8985"},{"content":{"text":[{"insert":"33","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuoh9q0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8987"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8988"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8989"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8990"},{"content":{"text":[{"insert":"32","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuokp90","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8992"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8993"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8994"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih8995"},{"content":{"text":[{"insert":"31","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuondt0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih8997"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih8998"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih8999"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89100"},{"content":{"text":[{"insert":"30","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuoq210","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89102"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89103"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89104"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89105"},{"content":{"text":[{"insert":"29","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuosis0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89107"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89108"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89109"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89110"},{"content":{"text":[{"insert":"28","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuov970","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89112"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89113"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89114"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89115"},{"content":{"text":[{"insert":"27","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpup1oc0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89117"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89118"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89119"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89120"},{"content":{"text":[{"insert":"26","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpup4sp0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89122"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89123"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89124"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89125"},{"content":{"text":[{"insert":"25","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpup7hl0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89127"},{"content":{"apple-news-follow_tagline":[{"insert":"Sign up to the TopGear Newsletter"}],"link":[{"insert":"SIGN UP","attributes":{"href":"https://www.topgear.com/newsletter-signup?utm_source=UK&utm_medium=homepage&utm_campaign=newsletter_signup"}}],"image":{"id":"72514","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"apple-news-follow","styles":{},"data":{"an-dark-mode-image":{"id":"72515"}},"id":"doc-1jhtp4fs91"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89128"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89129"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89130"},{"content":{"text":[{"insert":"24","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpupans0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89132"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89133"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89134"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89135"},{"content":{"text":[{"insert":"23","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpupdst0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89137"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89138"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89139"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89140"},{"content":{"text":[{"insert":"22","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx2"}]},"id":"doc-1jgpupgln0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89142"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89143"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89144"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89145"},{"content":{"text":[{"insert":"21","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpupjgs0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89147"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89148"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89149"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89150"},{"content":{"text":[{"insert":"20","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuplti0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89152"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89153"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89154"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89155"},{"content":{"text":[{"insert":"19","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpupojg0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89157"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89158"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89159"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89160"},{"content":{"text":[{"insert":"18","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuprqs0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89162"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89163"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89164"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89165"},{"content":{"text":[{"insert":"17","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpupuo30","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89167"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89168"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89169"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89170"},{"content":{"text":[{"insert":"16","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuq1vg0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89172"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89173"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89174"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89175"},{"content":{"text":[{"insert":"15","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuq4n00","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89177"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89178"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89179"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89180"},{"content":{"text":[{"insert":"14","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuq7i40","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89182"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89183"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89184"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89185"},{"content":{"text":[{"insert":"13","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuqaio0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89187"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1ih89188"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1ih89189"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1ih89190"},{"content":{"text":[{"insert":"12","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuqdap0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1ih89192"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa783"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e0"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e1"},{"content":{"text":[{"insert":"11","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuqfvd0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e3"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e4"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e5"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e6"},{"content":{"text":[{"insert":"10","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuqirn0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e8"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e9"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e10"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e11"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"9"},{"insert":" Xxxx"}]},"id":"doc-1jgpuqlmf0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e13"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e14"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e15"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e16"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"8"},{"insert":" Xxxx"}]},"id":"doc-1jgpuqoit0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e18"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e19"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e20"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e21"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"7"},{"insert":" Xxxx"}]},"id":"doc-1jgpuqr820","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e23"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e24"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e25"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e26"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"6"},{"insert":" Xxxx"}]},"id":"doc-1jgpuquui0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e28"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e29"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e30"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e31"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"5"},{"insert":" Xxxx"}]},"id":"doc-1jgpur1gb0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e33"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e34"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e35"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e36"},{"content":{"text":[{"insert":"4","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpur4j90","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e38"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e39"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e40"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e41"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"3"},{"insert":" Xxxx"}]},"id":"doc-1jgpur8l80","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e43"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e44"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e45"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e46"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"2"},{"insert":" Xxxx"}]},"id":"doc-1jgpurbdp0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e48"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcp1jb8e49"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e50"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcp1jb8e51"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"1"},{"insert":" Xxxx"}]},"id":"doc-1jgpure9f0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e53"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcp1jb8e54"},{"content":{},"identifier":"footer","styles":{"style":"_option1"},"id":"doc-1i1hesi1t7"},{"content":{"html":"<iframe allow=\"autoplay *; encrypted-media *; fullscreen *; clipboard-write\" frameborder=\"0\" height=\"450\" style=\"width:100%;max-width:660px;overflow:hidden;border-radius:10px;\" sandbox=\"allow-forms allow-popups allow-same-origin allow-scripts allow-storage-access-by-user-activation allow-top-navigation-by-user-activation\" src=\"https://embed.podcasts.apple.com/us/podcast/top-gear-magazine/id1614649540\"></iframe>"},"id":"doc-1jflpkntf0","identifier":"podcast","styles":{}},{"content":{},"identifier":"container","styles":{"style":"_option14"},"containers":{"main":[{"content":{"image":{"id":"48816","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content","inside-caption":"_caption-none","style":"_option7"},"id":"doc-1is9kes3t12"},{"content":{"title":[{"insert":"Discover more","attributes":{"color":"#79d6f2"}}]},"identifier":"subtitle","styles":{"text-align":"_align-middle"},"id":"doc-1is9kes3t13"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t14"},{"content":{"text":[{"insert":"Buy the magazine","attributes":{"href":"https://www.buysubscriptions.com/print/bbc-top-gear-magazine-subscription?promo=TGFMEV&utm_medium=brandsite&utm_source=topgear.com&utm_campaign=5for5_tgfmev&utm_content=footer-menu&style=brand","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t15"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t16"},{"content":{"text":[{"insert":"Visit our website","attributes":{"href":"https://www.topgear.com/","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t17"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t18"},{"content":{"text":[{"insert":"Listen to our podcast","attributes":{"href":"https://podcasts.apple.com/gb/podcast/top-gear-magazine/id1614649540","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t19"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t20"},{"content":{"text":[{"insert":"Watch films on YouTube","attributes":{"href":"https://www.youtube.com/user/TopGear","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t21"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t22"},{"content":{"text":[{"insert":"Sign up for our free newsletter","attributes":{"href":"https://www.topgear.com/newsletter-signup?utm_source=UK&utm_medium=homepage&utm_campaign=newsletter_signup","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t23"},{"content":{"text":[{"insert":"   "}]},"identifier":"body","styles":{},"id":"doc-1is9kes3t24"},{"content":{},"identifier":"container","styles":{"style":"_option15"},"containers":{"main":[{"content":{"text":[{"insert":"A","attributes":{"href":"https://www.facebook.com/topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t26"},{"content":{"text":[{"insert":"C","attributes":{"href":"https://x.com/bbc_topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t27"},{"content":{"text":[{"insert":"D","attributes":{"href":"https://www.youtube.com/user/TopGear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t28"},{"content":{"text":[{"insert":"E","attributes":{"href":"https://www.instagram.com/topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t29"},{"content":{"text":[{"insert":"F","attributes":{"color":"#102935","href":"https://www.tiktok.com/@topgear?is_from_webapp=1&sender_device=pc"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t30"}]},"inlineStyles":{"background-color":"rgb(121, 214, 242)"},"data":{"an-content-display":"collection"},"id":"doc-1is9kes3t25"}]},"inlineStyles":{"background-color":"rgb(0, 39, 48)"},"id":"doc-1is9kes3t11"}]}},
  ascending: {"version":"2.4","data":{"content":[{"content":{},"identifier":"header-image","styles":{},"data":{"an-image-role":"photo"},"id":"doc-1jj1hv1ft0"},{"content":{},"id":"doc-1jgpv33pe0","identifier":"title","styles":{"text-align":"_align-middle"}},{"content":{"title":[]},"identifier":"subtitle","styles":{"text-align":"_align-middle"},"id":"doc-1i1hesi1t2"},{"content":{"name":[]},"identifier":"author","styles":{"text-align":"_align-middle"},"id":"doc-1i1hesi1t3"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa780"},{"content":{"text":[{"attributes":{"color":"#ff3f63"},"insert":"01"},{"insert":" Xxxx"}]},"id":"doc-1jgpuu24b0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa782"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa783"},{"content":{"apple-news-follow_tagline":[{"insert":"For more content follow this channel"}],"link":[{"insert":"FOLLOW","attributes":{"href":"https://apple.news/TRtxUc-x1TtSqnoPE8MyWFA?subscribe=1"}}],"image":{"id":"72519","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"apple-news-follow","styles":{},"data":{"an-dark-mode-image":{"id":"72518"}},"id":"doc-1jhtpfbqq0"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa784"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa785"},{"content":{"text":[{"insert":"02","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuu7hg0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa787"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa788"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa789"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7810"},{"content":{"text":[{"insert":"03","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuua8n0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7812"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7813"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7814"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7815"},{"content":{"text":[{"insert":"04","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuuf0a0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7817"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7818"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7819"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7820"},{"content":{"text":[{"insert":"05","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuuhlf0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7822"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7823"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7824"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7825"},{"content":{"text":[{"insert":"06","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuuk360","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7827"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7828"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7829"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7830"},{"content":{"text":[{"insert":"07","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuumm00","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7832"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7833"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7834"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7835"},{"content":{"text":[{"insert":"08","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuupru0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7837"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7838"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7839"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7840"},{"content":{"text":[{"insert":"09","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuusi00","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7842"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7843"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7844"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7845"},{"content":{"text":[{"insert":"10","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuuv0t0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7847"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7848"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7849"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7850"},{"content":{"text":[{"insert":"11","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuv1jn0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7852"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7853"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7854"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7855"},{"content":{"text":[{"insert":"12","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuv4350","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7857"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7858"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7859"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7860"},{"content":{"text":[{"insert":"13","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuv8o80","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7862"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7863"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7864"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7865"},{"content":{"text":[{"insert":"14","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvbdd0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7867"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7868"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7869"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7870"},{"content":{"text":[{"insert":"15","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvdul0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7872"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7873"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7874"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7875"},{"content":{"text":[{"insert":"16","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvgfk0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7877"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7878"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7879"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7880"},{"content":{"text":[{"insert":"17","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvj6s0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7882"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7883"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7884"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7885"},{"content":{"text":[{"insert":"18","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvlns0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7887"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7888"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7889"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7890"},{"content":{"text":[{"insert":"19","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvovp0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7892"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7893"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7894"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa7895"},{"content":{"text":[{"insert":"20","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvrfr0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa7897"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa7898"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa7899"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78100"},{"content":{"text":[{"insert":"21","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpuvu180","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78102"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78103"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78104"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78105"},{"content":{"text":[{"insert":"22","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv013d0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78107"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78108"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78109"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78110"},{"content":{"text":[{"insert":"23","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv05n20","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78112"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78113"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78114"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78115"},{"content":{"text":[{"insert":"24","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv08bi0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78117"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78118"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78119"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78120"},{"content":{"text":[{"insert":"25","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0avu0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78122"},{"content":{"apple-news-follow_tagline":[{"insert":"Sign up to the TopGear Newsletter"}],"link":[{"insert":"SIGN UP","attributes":{"href":"https://www.topgear.com/newsletter-signup?utm_source=UK&utm_medium=homepage&utm_campaign=newsletter_signup"}}],"image":{"id":"72517","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"apple-news-follow","styles":{},"data":{"an-dark-mode-image":{"id":"72520"}},"id":"doc-1jhtpfbqq1"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78123"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78124"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78125"},{"content":{"text":[{"insert":"26","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0dgu0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78127"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78128"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78129"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78130"},{"content":{"text":[{"insert":"27","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0ha50","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78132"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78133"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78134"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78135"},{"content":{"text":[{"insert":"28","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0k6j0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78137"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78138"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78139"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcousa78140"},{"content":{"text":[{"insert":"29","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0mrk0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcousa78142"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcousa78143"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcousa78144"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q0"},{"content":{"text":[{"insert":"30","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0pjd0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q2"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q3"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q4"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q5"},{"content":{"text":[{"insert":"31","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0s370","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q7"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q8"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q9"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q10"},{"content":{"text":[{"insert":"32","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv0upl0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q12"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q13"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q14"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q15"},{"content":{"text":[{"insert":"33","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv11970","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q17"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q18"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q19"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q20"},{"content":{"text":[{"insert":"34","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv140f0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q22"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q23"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q24"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q25"},{"content":{"text":[{"insert":"35","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv16ls0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q27"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q28"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q29"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q30"},{"content":{"text":[{"insert":"36","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv1ec70","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q32"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q33"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q34"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q35"},{"content":{"text":[{"insert":"37","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv1gqg0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q37"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q38"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q39"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q40"},{"content":{"text":[{"insert":"38","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv1jft0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q42"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q43"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q44"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q45"},{"content":{"text":[{"insert":"39","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv1m190","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q47"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q52"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q53"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q54"},{"content":{"text":[{"insert":"40","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv1otj0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q56"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q57"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q58"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q59"},{"content":{"text":[{"insert":"41","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv1r5o0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q61"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q62"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q63"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q64"},{"content":{"text":[{"insert":"42","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv1vbn0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q66"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q67"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q68"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q69"},{"content":{"text":[{"insert":"43","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv24120","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q71"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q72"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q73"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q74"},{"content":{"text":[{"insert":"44","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv26h00","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q76"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q77"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q78"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q79"},{"content":{"text":[{"insert":"45","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv29lq0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q81"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q82"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q83"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q84"},{"content":{"text":[{"insert":"46","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv2c9t0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q86"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q87"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q88"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q89"},{"content":{"text":[{"insert":"47","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv2emn0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q91"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q92"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q93"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q94"},{"content":{"text":[{"insert":"48","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv2h5k0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q96"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q97"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q98"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q99"},{"content":{"text":[{"insert":"49","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv2jo60","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q101"},{"content":{"text":[{"insert":"  "}]},"identifier":"body","styles":{},"id":"doc-1jcouru7q102"},{"content":{},"identifier":"separator","styles":{"style":"_option1"},"id":"doc-1jcouru7q103"},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1jcouru7q104"},{"content":{"text":[{"insert":"50","attributes":{"color":"#ff3f63"}},{"insert":" Xxxx"}]},"id":"doc-1jgpv2ovt0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jcouru7q106"},{"content":{},"identifier":"footer","styles":{"style":"_option1"},"id":"doc-1i1hesi1t7"},{"content":{"html":"<iframe allow=\"autoplay *; encrypted-media *; fullscreen *; clipboard-write\" frameborder=\"0\" height=\"450\" style=\"width:100%;max-width:660px;overflow:hidden;border-radius:10px;\" sandbox=\"allow-forms allow-popups allow-same-origin allow-scripts allow-storage-access-by-user-activation allow-top-navigation-by-user-activation\" src=\"https://embed.podcasts.apple.com/us/podcast/top-gear-magazine/id1614649540\"></iframe>"},"id":"doc-1jflpilaj0","identifier":"podcast","styles":{}},{"content":{},"identifier":"container","styles":{"style":"_option14"},"containers":{"main":[{"content":{"image":{"id":"48816","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content","inside-caption":"_caption-none","style":"_option7"},"id":"doc-1is9kes3t12"},{"content":{"title":[{"insert":"Discover more","attributes":{"color":"#79d6f2"}}]},"identifier":"subtitle","styles":{"text-align":"_align-middle"},"id":"doc-1is9kes3t13"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t14"},{"content":{"text":[{"insert":"Buy the magazine","attributes":{"href":"https://www.buysubscriptions.com/print/bbc-top-gear-magazine-subscription?promo=TGFMEV&utm_medium=brandsite&utm_source=topgear.com&utm_campaign=5for5_tgfmev&utm_content=footer-menu&style=brand","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t15"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t16"},{"content":{"text":[{"insert":"Visit our website","attributes":{"href":"https://www.topgear.com/","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t17"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t18"},{"content":{"text":[{"insert":"Listen to our podcast","attributes":{"href":"https://podcasts.apple.com/gb/podcast/top-gear-magazine/id1614649540","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t19"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t20"},{"content":{"text":[{"insert":"Watch films on YouTube","attributes":{"href":"https://www.youtube.com/user/TopGear","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t21"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t22"},{"content":{"text":[{"insert":"Sign up for our free newsletter","attributes":{"href":"https://www.topgear.com/newsletter-signup?utm_source=UK&utm_medium=homepage&utm_campaign=newsletter_signup","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t23"},{"content":{"text":[{"insert":"   "}]},"identifier":"body","styles":{},"id":"doc-1is9kes3t24"},{"content":{},"identifier":"container","styles":{"style":"_option15"},"containers":{"main":[{"content":{"text":[{"insert":"A","attributes":{"href":"https://www.facebook.com/topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t26"},{"content":{"text":[{"insert":"C","attributes":{"href":"https://x.com/bbc_topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t27"},{"content":{"text":[{"insert":"D","attributes":{"href":"https://www.youtube.com/user/TopGear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t28"},{"content":{"text":[{"insert":"E","attributes":{"href":"https://www.instagram.com/topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t29"},{"content":{"text":[{"insert":"F","attributes":{"color":"#102935","href":"https://www.tiktok.com/@topgear?is_from_webapp=1&sender_device=pc"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t30"}]},"inlineStyles":{"background-color":"rgb(121, 214, 242)"},"data":{"an-content-display":"collection"},"id":"doc-1is9kes3t25"}]},"inlineStyles":{"background-color":"rgb(0, 39, 48)"},"id":"doc-1is9kes3t11"}]}},
  crosshead: {"version":"2.4","data":{"content":[{"content":{},"id":"doc-1jj1hqqba0","identifier":"header-image","styles":{},"data":{"an-image-role":"photo"}},{"content":{},"id":"doc-1jblfrog80","identifier":"title","styles":{"text-align":"_align-middle"}},{"content":{"title":[]},"identifier":"subtitle","styles":{"text-align":"_align-middle"},"id":"doc-1i1hesi1t2"},{"content":{"name":[]},"identifier":"author","styles":{"text-align":"_align-middle"},"id":"doc-1i1hesi1t3"},{"content":{},"id":"doc-1jflpepua0","identifier":"body","styles":{"style":"_option8"}},{"content":{"text":[{"insert":"[[Main Style Crosshead]]"}]},"id":"doc-1jgpu5gmr0","identifier":"title","styles":{"style":"_option2"}},{"content":{"text":[{"insert":"[[Alternate style Crosshead]]"}]},"id":"doc-1jblfshc20","identifier":"crosshead","styles":{}},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1i1hesi1t4"},{"content":{"apple-news-follow_tagline":[{"insert":"For more content follow this channel"}],"link":[{"insert":"FOLLOW","attributes":{"href":"https://apple.news/TRtxUc-x1TtSqnoPE8MyWFA?subscribe=1"}}],"image":{"id":"72512","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"apple-news-follow","styles":{},"data":{"an-dark-mode-image":{"id":"72509"}},"id":"doc-1jhtp4ak40"},{"content":{"text":[]},"identifier":"title","styles":{"style":"_option2"},"id":"doc-1jj1hs2gn0"},{"content":{"text":[]},"identifier":"crosshead","styles":{},"id":"doc-1jj1hs2gn1"},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jj1hs2gn2"},{"content":{"text":[]},"identifier":"title","styles":{"style":"_option2"},"id":"doc-1jj1hs4mm0"},{"content":{"text":[]},"identifier":"crosshead","styles":{},"id":"doc-1jj1hs4mm1"},{"content":{"text":[]},"identifier":"body","styles":{"style":"_option1"},"id":"doc-1jj1hs4mm2"},{"content":{"text":[{"insert":"  "}]},"id":"doc-1jblfsl5c0","identifier":"body","styles":{}},{"content":{"text":[{"insert":"SCORE: "},{"insert":"X/10","attributes":{"color":"#ff3f63"}}]},"id":"doc-1jhqibg5l0","identifier":"title","styles":{"style":"_option3"}},{"content":{},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content"},"data":{"an-image-role":"photo"},"id":"doc-1i1hesi1t5"},{"content":{},"identifier":"footer","styles":{"style":"_option1"},"id":"doc-1i1hesi1t7"},{"content":{"apple-news-follow_tagline":[{"insert":"Sign up to the TopGear Newsletter"}],"link":[{"insert":"SIGN UP","attributes":{"href":"https://www.topgear.com/newsletter-signup?utm_source=UK&utm_medium=homepage&utm_campaign=newsletter_signup"}}],"image":{"id":"72510","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"apple-news-follow","styles":{},"data":{"an-dark-mode-image":{"id":"72511"}},"id":"doc-1jhtp4ak41"},{"content":{"html":"<iframe allow=\"autoplay *; encrypted-media *; fullscreen *; clipboard-write\" frameborder=\"0\" height=\"450\" style=\"width:100%;max-width:660px;overflow:hidden;border-radius:10px;\" sandbox=\"allow-forms allow-popups allow-same-origin allow-scripts allow-storage-access-by-user-activation allow-top-navigation-by-user-activation\" src=\"https://embed.podcasts.apple.com/us/podcast/top-gear-magazine/id1614649540\"></iframe>"},"id":"doc-1jflpfeof0","identifier":"podcast","styles":{}},{"content":{},"identifier":"container","styles":{"style":"_option14"},"containers":{"main":[{"content":{"image":{"id":"48816","focuspoint":{"x":0.5,"y":0.5},"cropper":false}},"identifier":"image","styles":{"fitting":"_fit-frame-height-to-content","inside-caption":"_caption-none","style":"_option7"},"id":"doc-1is9kes3t12"},{"content":{"title":[{"insert":"Discover more","attributes":{"color":"#79d6f2"}}]},"identifier":"subtitle","styles":{"text-align":"_align-middle"},"id":"doc-1is9kes3t13"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t14"},{"content":{"text":[{"insert":"Buy the magazine","attributes":{"href":"https://www.buysubscriptions.com/print/bbc-top-gear-magazine-subscription?promo=TGFMEV&utm_medium=brandsite&utm_source=topgear.com&utm_campaign=5for5_tgfmev&utm_content=footer-menu&style=brand","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t15"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t16"},{"content":{"text":[{"insert":"Visit our website","attributes":{"href":"https://www.topgear.com/","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t17"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t18"},{"content":{"text":[{"insert":"Listen to our podcast","attributes":{"href":"https://podcasts.apple.com/gb/podcast/top-gear-magazine/id1614649540","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t19"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t20"},{"content":{"text":[{"insert":"Watch films on YouTube","attributes":{"href":"https://www.youtube.com/user/TopGear","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t21"},{"content":{},"identifier":"separator","styles":{"style":"_option7"},"id":"doc-1is9kes3t22"},{"content":{"text":[{"insert":"Sign up for our free newsletter","attributes":{"href":"https://www.topgear.com/newsletter-signup?utm_source=UK&utm_medium=homepage&utm_campaign=newsletter_signup","color":"#ffffff"}}]},"identifier":"body","styles":{"style":"_option7","text-align":"_align-middle"},"id":"doc-1is9kes3t23"},{"content":{"text":[{"insert":"   "}]},"identifier":"body","styles":{},"id":"doc-1is9kes3t24"},{"content":{},"identifier":"container","styles":{"style":"_option15"},"containers":{"main":[{"content":{"text":[{"insert":"A","attributes":{"href":"https://www.facebook.com/topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t26"},{"content":{"text":[{"insert":"C","attributes":{"href":"https://x.com/bbc_topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t27"},{"content":{"text":[{"insert":"D","attributes":{"href":"https://www.youtube.com/user/TopGear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t28"},{"content":{"text":[{"insert":"E","attributes":{"href":"https://www.instagram.com/topgear","color":"#102935"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t29"},{"content":{"text":[{"insert":"F","attributes":{"color":"#102935","href":"https://www.tiktok.com/@topgear?is_from_webapp=1&sender_device=pc"}}]},"identifier":"crosshead","styles":{"style":"_option6"},"id":"doc-1is9kes3t30"}]},"inlineStyles":{"background-color":"rgb(121, 214, 242)"},"data":{"an-content-display":"collection"},"id":"doc-1is9kes3t25"}]},"inlineStyles":{"background-color":"rgb(0, 39, 48)"},"id":"doc-1is9kes3t11"}]}}
};

// ─── Utilities ────────────────────────────────────────────────────────────
const genId = () => 'doc-' + Math.random().toString(36).slice(2, 12);
const deepClone = o => JSON.parse(JSON.stringify(o));

function setArticleTitle(header, text, deltas) {
  const c = header.find(c => c.identifier === 'title');
  if (c) c.content = { text: deltas && deltas.length ? deepClone(deltas) : (text ? [{ insert: text }] : []) };
}
function setSubtitle(header, text, deltas) {
  const c = header.find(c => c.identifier === 'subtitle');
  if (c) c.content = { title: deltas && deltas.length ? deepClone(deltas) : (text ? [{ insert: text }] : []) };
}
function setAuthor(header, name) {
  const c = header.find(c => c.identifier === 'author');
  if (c) {
    c.content = name
      ? { name: [{ insert: 'Words:', attributes: { bold: true } }, { insert: ' ' + name }] }
      : { name: [] };
  }
}

// Converts DOM nodes to deltas, preserving bold, italic and links through nesting.
// opts.italicOnly: for headline contexts (titles, crossheads, metadata values) —
// those components are already bold by template style, so only italic/links carry over.
function nodesToDeltas(nodes, opts = {}) {
  const ops = [];
  function walk(node, fmt) {
    if (node.nodeType === Node.TEXT_NODE) {
      if (!node.textContent) return;
      const attrs = {};
      if (fmt.bold && !opts.italicOnly) attrs.bold = true;
      if (fmt.italic) attrs.italic = true;
      if (fmt.href) attrs.href = fmt.href;
      ops.push(Object.keys(attrs).length
        ? { insert: node.textContent, attributes: attrs }
        : { insert: node.textContent });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const tag = node.tagName.toLowerCase();
    const next = Object.assign({}, fmt);
    if (tag === 'strong' || tag === 'b') next.bold = true;
    if (tag === 'em' || tag === 'i') next.italic = true;
    if (tag === 'a' && node.getAttribute('href')) next.href = node.getAttribute('href');
    for (const child of node.childNodes) walk(child, next);
  }
  for (const node of nodes) walk(node, {});
  return ops;
}

function paraToDeltas(para, opts) {
  return nodesToDeltas(para.childNodes, opts);
}

// Trim outer whitespace of a delta run without touching inner spacing
function trimDeltas(ops) {
  const out = ops.map(o => Object.assign({}, o)).filter(o => o.insert);
  while (out.length && !out[0].insert.trim()) out.shift();
  if (out.length) out[0].insert = out[0].insert.replace(/^\s+/, '');
  while (out.length && !out[out.length - 1].insert.trim()) out.pop();
  if (out.length) out[out.length - 1].insert = out[out.length - 1].insert.replace(/\s+$/, '');
  return out;
}

// Deltas for the text that follows `prefix` in a paragraph (e.g. a metadata value)
function deltasAfterText(ops, prefix) {
  const full = ops.map(o => o.insert).join('');
  const idx = full.toLowerCase().indexOf(prefix.toLowerCase());
  if (idx === -1) return null;
  let skip = idx + prefix.length;
  const out = [];
  for (const op of ops) {
    if (skip >= op.insert.length) { skip -= op.insert.length; continue; }
    out.push(Object.assign({}, op, { insert: op.insert.slice(skip) }));
    skip = 0;
  }
  return trimDeltas(out);
}

function joinDeltas(parts) {
  const result = [];
  parts.forEach((part, i) => {
    if (i > 0) result.push({ insert: '\n\n' });
    result.push(...part);
  });
  return result;
}

// ─── Docx parsing ────────────────────────────────────────────────────────
const METADATA_PREFIXES = [
  { key: 'feedHeadline', prefix: 'Feed headline:' },
  { key: 'title',        prefix: 'Article headline:' },
  { key: 'subtitle',     prefix: 'Article subhead:' },
  { key: 'author',       prefix: 'Words:' },
];

function matchMetaPrefix(text) {
  for (const { key, prefix } of METADATA_PREFIXES) {
    if (text.toLowerCase().startsWith(prefix.toLowerCase())) {
      return { key, prefix, value: text.slice(prefix.length).trim() };
    }
  }
  return null;
}

// Some docs style crossheads/metadata as Word headings rather than bold paragraphs,
// so headings must be walked too — mammoth emits them as <h1>–<h6>.
function paras(html) {
  const doc = new DOMParser().parseFromString('<div>' + html + '</div>', 'text/html');
  return Array.from(doc.querySelectorAll('p, h1, h2, h3, h4, h5, h6'));
}

// Lines that are editor instructions, not article content
const SKIP_PREFIXES = ['pics:', 'web gallery', 'imbed ', 'imbedded', 'embed ', 'embargo'];
function shouldSkip(text) {
  const lower = text.toLowerCase();
  return SKIP_PREFIXES.some(p => lower.startsWith(p)) || /^https?:\/\/\S+$/.test(text);
}

// Returns { crosshead, remainingBody } if paragraph is (or starts with) bold, else null.
// Handles:
//   - Fully bold paragraph  →  crosshead = full text, remainingBody = ''
//   - Paragraph starting with bold  →  crosshead = bold part, remainingBody = rest
function extractCrosshead(para) {
  const text = para.textContent.trim();
  if (!text) return null;
  const headlineDeltas = () => trimDeltas(paraToDeltas(para, { italicOnly: true }));
  const firstStrong = para.querySelector('strong');
  if (!firstStrong) {
    // Heading elements are crossheads even without explicit bold runs
    if (/^H[1-6]$/.test(para.tagName)) {
      return { crosshead: text, crossheadDeltas: headlineDeltas(), remainingBody: '', remainingDeltas: [] };
    }
    return null;
  }
  const strongText = firstStrong.textContent.trim();
  if (!strongText) return null;
  // Fully bold
  if (strongText === text) {
    return { crosshead: text, crossheadDeltas: headlineDeltas(), remainingBody: '', remainingDeltas: [] };
  }
  // Starts with bold (mixed paragraph: bold question + plain answer)
  const firstEl = para.firstElementChild;
  if (firstEl && firstEl.tagName === 'STRONG') {
    const rest = [];
    let past = false;
    for (const child of para.childNodes) {
      if (child === firstEl) { past = true; continue; }
      if (past) rest.push(child);
    }
    return {
      crosshead: strongText,
      crossheadDeltas: trimDeltas(nodesToDeltas([firstStrong], { italicOnly: true })),
      remainingBody: text.slice(strongText.length).trim(),
      remainingDeltas: nodesToDeltas(rest),
    };
  }
  return null;
}

// Rich version of a numbered-entry name ("12. Route 66" minus the "12. ")
function entryNameDeltas(p) {
  const ops = paraToDeltas(p, { italicOnly: true });
  const m = ops.map(o => o.insert).join('').match(/^\s*\d+\.\s+/);
  return m ? deltasAfterText(ops, m[0]) : null;
}

function parseNumbered(html) {
  const ENTRY_RE = /^(\d+)\.\s+(.+)$/;
  const meta = { feedHeadline: '', title: '', subtitle: '', author: '', unrecognized: [], deltas: {} };
  const entries = [];
  let inEntries = false;
  let pending = null;
  let pendingMetaKey = null;

  for (const p of paras(html)) {
    const text = p.textContent.trim();
    if (!text) continue;

    if (!inEntries) {
      // Multi-line metadata: previous line was a label with no value
      if (pendingMetaKey) {
        if (!shouldSkip(text) && !ENTRY_RE.test(text)) {
          meta[pendingMetaKey] = text;
          meta.deltas[pendingMetaKey] = trimDeltas(paraToDeltas(p, { italicOnly: true }));
          pendingMetaKey = null;
          continue;
        }
        pendingMetaKey = null;
      }

      const m = text.match(ENTRY_RE);
      if (m) {
        inEntries = true;
        pending = { number: parseInt(m[1]), name: m[2], nameDeltas: entryNameDeltas(p), bodyParts: [] };
        continue;
      }
      const match = matchMetaPrefix(text);
      if (match) {
        if (match.value === '') { pendingMetaKey = match.key; }
        else {
          meta[match.key] = match.value;
          meta.deltas[match.key] = deltasAfterText(paraToDeltas(p, { italicOnly: true }), match.prefix);
        }
      } else if (!shouldSkip(text)) {
        meta.unrecognized.push(text);
      }
    } else {
      const m = text.match(ENTRY_RE);
      if (m) {
        if (pending) entries.push(pending);
        pending = { number: parseInt(m[1]), name: m[2], nameDeltas: entryNameDeltas(p), bodyParts: [] };
      } else {
        if (pending) {
          const deltas = paraToDeltas(p);
          if (deltas.length) pending.bodyParts.push(deltas);
        }
      }
    }
  }
  if (pending) entries.push(pending);
  return { meta, entries };
}

function parseCrosshead(html) {
  const meta = { feedHeadline: '', title: '', subtitle: '', author: '', intro: '', score: '', unrecognized: [], deltas: {}, introParts: [] };
  const entries = [];
  let inContent = false;
  let curCrosshead = null;
  let curCrossheadDeltas = null;
  let curBodyParts = [];
  let pendingMetaKey = null;

  for (const p of paras(html)) {
    const text = p.textContent.trim();
    if (!text) continue;

    // 1. Multi-line metadata: previous line was a label with no value
    if (pendingMetaKey) {
      if (!shouldSkip(text)) {
        // Don't consume the value if it looks like a crosshead we've been waiting for
        const crosshead = extractCrosshead(p);
        const isMeta = !!matchMetaPrefix(text);
        if (!isMeta) {
          meta[pendingMetaKey] = text;
          meta.deltas[pendingMetaKey] = trimDeltas(paraToDeltas(p, { italicOnly: true }));
          pendingMetaKey = null;
          continue;
        }
      }
      pendingMetaKey = null;
    }

    // 2. Always check metadata FIRST — even bold lines can be metadata (e.g. "Words: Jack Rix")
    const metaMatch = matchMetaPrefix(text);
    if (metaMatch) {
      if (metaMatch.value === '') { pendingMetaKey = metaMatch.key; }
      else {
        meta[metaMatch.key] = metaMatch.value;
        meta.deltas[metaMatch.key] = deltasAfterText(paraToDeltas(p, { italicOnly: true }), metaMatch.prefix);
      }
      continue;
    }

    // 3. Skip editor instructions and bare URLs
    if (shouldSkip(text)) continue;

    // 3.5. Detect score line (e.g. "SCORE: 7/10") — must run before crosshead detection
    if (/^score:\s*\d/i.test(text)) {
      meta.score = text;
      continue;
    }

    // 4. Detect crosshead (fully bold, or starts with bold)
    const crosshead = extractCrosshead(p);
    if (crosshead) {
      if (inContent && curCrosshead !== null) {
        entries.push({ crosshead: curCrosshead, crossheadDeltas: curCrossheadDeltas, bodyParts: curBodyParts });
      }
      inContent = true;
      curCrosshead = crosshead.crosshead;
      curCrossheadDeltas = crosshead.crossheadDeltas;
      curBodyParts = crosshead.remainingDeltas && crosshead.remainingDeltas.length
        ? [crosshead.remainingDeltas]
        : [];
    } else {
      // Plain body text
      if (inContent) {
        if (curCrosshead !== null) {
          const deltas = paraToDeltas(p);
          if (deltas.length) curBodyParts.push(deltas);
        }
      } else {
        meta.intro = meta.intro ? meta.intro + '\n\n' + text : text;
        meta.introParts.push(trimDeltas(paraToDeltas(p)));
      }
    }
  }
  if (curCrosshead !== null) entries.push({ crosshead: curCrosshead, crossheadDeltas: curCrossheadDeltas, bodyParts: curBodyParts });
  return { meta, entries };
}

// ─── Template building ───────────────────────────────────────────────────
function buildNumbered(template, meta, entries, type) {
  const c = deepClone(template).data.content;

  // Header: first 4 components [header-image, title, subtitle, author]
  const header = c.slice(0, 4);
  setArticleTitle(header, meta.title, meta.titleDeltas);
  setSubtitle(header, meta.subtitle, meta.subtitleDeltas);
  setAuthor(header, meta.author);
  header.forEach(comp => { comp.id = genId(); });

  // Footer: everything from the 'footer' component onwards
  const footerIdx = c.findIndex(comp => comp.identifier === 'footer');
  const footer = deepClone(c.slice(footerIdx));
  footer.forEach(comp => { comp.id = genId(); });

  // Channel apple-news-follow: first one in the first entry block (indices 4–12)
  const channelFollow = deepClone(c.slice(4, 13).find(comp => comp.identifier === 'apple-news-follow'));

  // Canonical entry: second entry in the template (clean, no apple-news-follow)
  // Find second 'image' component (first entry starts at 4, second entry starts after first separator)
  const firstSepIdx = c.findIndex((comp, i) => i > 4 && comp.identifier === 'separator');
  const secondEntryStart = firstSepIdx + 1;
  const secondSepIdx = c.findIndex((comp, i) => i > secondEntryStart && comp.identifier === 'separator');
  const canonical = deepClone(c.slice(secondEntryStart, secondSepIdx + 1));
  // canonical = [image, title, body(_option1), body(spacer), separator]

  // For type 1 countdown: apple goes after body (index 2 in canonical), splice at 3
  // For type 2 ascending: apple goes after spacer (index 3 in canonical), splice at 4
  const appleInsertAt = type === 'countdown' ? 3 : 4;

  // Sort entries
  entries.sort((a, b) => type === 'countdown' ? b.number - a.number : a.number - b.number);

  const result = [...header];

  entries.forEach((entry, i) => {
    const group = deepClone(canonical);

    // Update title component (keep the template's coloured number op; NBSP separates number and name)
    const titleComp = group.find(comp => comp.identifier === 'title');
    if (titleComp && titleComp.content && titleComp.content.text && titleComp.content.text.length >= 2) {
      const numberOp = titleComp.content.text[0];
      numberOp.insert = type === 'ascending'
        ? String(entry.number).padStart(2, '0')
        : String(entry.number);
      titleComp.content.text = entry.nameDeltas && entry.nameDeltas.length
        ? [numberOp, { insert: '\u00A0' }, ...deepClone(entry.nameDeltas)]
        : [numberOp, { insert: '\u00A0' + entry.name }];
    }

    // Update body (description = has style _option1)
    const bodyCompIdx = group.findIndex(comp =>
      comp.identifier === 'body' && comp.styles && comp.styles.style === '_option1'
    );
    const bodyComp = group[bodyCompIdx];
    if (bodyComp) bodyComp.content = { text: entry.bodyParts[0] || [] };

    // Fresh IDs
    group.forEach(comp => { comp.id = genId(); });

    // Splice in extra body paragraphs (plain styles) after the first body
    const extraBodies = (entry.bodyParts || []).slice(1).map(deltas => ({
      identifier: 'body', styles: {}, content: { text: deltas }, id: genId()
    }));
    group.splice(bodyCompIdx + 1, 0, ...extraBodies);

    // Insert apple-news-follow after first entry (offset by extra paragraphs)
    if (i === 0 && channelFollow) {
      const af = deepClone(channelFollow);
      af.id = genId();
      group.splice(appleInsertAt + extraBodies.length, 0, af);
    }

    result.push(...group);
  });

  result.push(...footer);
  return { version: template.version, data: { content: result } };
}

function buildCrosshead(template, meta, entries) {
  const c = deepClone(template).data.content;

  const header = deepClone(c.slice(0, 4));
  setArticleTitle(header, meta.title, meta.titleDeltas);
  setSubtitle(header, meta.subtitle, meta.subtitleDeltas);
  setAuthor(header, meta.author);
  header.forEach(comp => { comp.id = genId(); });

  // Articles that open straight on a crosshead have no intro — drop the component
  const introBody = meta.intro ? deepClone(c[4]) : null;
  if (introBody) {
    introBody.id = genId();
    introBody.content = {
      text: meta.introParts && meta.introParts.length
        ? joinDeltas(meta.introParts)
        : [{ insert: meta.intro }]
    };
  }

  const channelFollow = deepClone(c[8]);
  channelFollow.id = genId();

  // Canonical section = indices 9–11: [title, crosshead, body]
  const canonical = deepClone(c.slice(9, 12));

  // Fixed tail: spacer(15), score-title(16), image(17), footer onwards(18+)
  const spacer = deepClone(c[15]);
  const scoreTitle = deepClone(c[16]);
  const reviewImage = deepClone(c[17]);
  const footerComps = deepClone(c.slice(18));

  const result = introBody ? [...header, introBody] : [...header];

  entries.forEach((entry, i) => {
    // Filter out the crosshead component — questions go in title(_option2) per reference structure
    const group = deepClone(canonical).filter(comp => comp.identifier !== 'crosshead');

    const titleComp = group.find(comp => comp.identifier === 'title');
    if (titleComp) {
      titleComp.content = {
        text: entry.crossheadDeltas && entry.crossheadDeltas.length
          ? deepClone(entry.crossheadDeltas)
          : (entry.crosshead ? [{ insert: entry.crosshead }] : [])
      };
    }

    const bodyComp = group.find(comp => comp.identifier === 'body');
    if (bodyComp) bodyComp.content = { text: entry.bodyParts[0] || [] };

    group.forEach(comp => { comp.id = genId(); });

    const extraBodies = (entry.bodyParts || []).slice(1).map(deltas => ({
      identifier: 'body', styles: {}, content: { text: deltas }, id: genId()
    }));

    result.push(...group, ...extraBodies);

    if (i === 0) result.push(deepClone(channelFollow));
  });

  if (meta.score) scoreTitle.content = { text: [{ insert: meta.score }] };
  [spacer, scoreTitle, reviewImage].forEach(comp => {
    comp.id = genId();
    result.push(comp);
  });
  footerComps.forEach(comp => { comp.id = genId(); result.push(comp); });

  return { version: template.version, data: { content: result } };
}


  // ─── End conversion engine ────────────────────────────────────────────────

  var DIGITAL_MIME = 'application/ww-digital+json';

  // Word parsing dependency, loaded on demand and kept plugin-local per the
  // SDK guidance on managing external dependencies.
  var mammothPromise = null;
  function loadMammoth() {
    if (window.mammoth) return Promise.resolve(window.mammoth);
    if (!mammothPromise) {
      mammothPromise = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js';
        s.onload = function () { resolve(window.mammoth); };
        s.onerror = function () {
          mammothPromise = null;
          reject(new Error('Could not load the Word parser (mammoth.js). Check that cdn.jsdelivr.net is reachable.'));
        };
        document.head.appendChild(s);
      });
    }
    return mammothPromise;
  }

  // ─── Studio Server API (same origin — session comes from ticket or cookie) ─
  function getTicket() {
    try {
      var info = ContentStationSdk.getInfo();
      return (info && info.Ticket) || '';
    } catch (e) { return ''; }
  }

  function serverIndexUrl() {
    var rel = (window.csConfig && window.csConfig.serverUrl) || '../server/index.php';
    return new URL(rel, window.location.href).href;
  }

  function transferUrl() {
    try {
      var fs = ContentStationSdk.getInfo().ServerInfo.FeatureSet || [];
      for (var i = 0; i < fs.length; i++) {
        if (fs[i].Key === 'FileUploadUrl' && fs[i].Value) return fs[i].Value;
      }
    } catch (e) { /* fall through */ }
    return serverIndexUrl().replace(/index\.php.*$/, 'transferindex.php');
  }

  function callServer(method, params) {
    return fetch(serverIndexUrl() + '?protocol=JSON', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ method: method, id: '1', params: [params], jsonrpc: '2.0' }),
    }).then(function (r) {
      if (!r.ok) throw new Error(method + ' failed: HTTP ' + r.status);
      return r.json();
    }).then(function (j) {
      if (j.error) throw new Error(method + ' failed: ' + ((j.error.data && j.error.data.detail) || j.error.message || JSON.stringify(j.error)));
      return j.result;
    });
  }

  function uploadToTransferServer(content) {
    var ticket = getTicket();
    var url = transferUrl() + '?ticket=' + encodeURIComponent(ticket) + '&uploadtokens=1';
    return fetch(url, { credentials: 'same-origin' })
      .then(function (r) {
        if (!r.ok) throw new Error('Transfer Server token request failed: HTTP ' + r.status);
        return r.json();
      })
      .then(function (tokens) {
        var tokenUrl = Array.isArray(tokens) ? tokens[0] : tokens;
        if (!tokenUrl) throw new Error('Transfer Server returned no upload token');
        return fetch(tokenUrl, {
          method: 'PUT',
          credentials: 'same-origin',
          headers: { 'Content-Type': DIGITAL_MIME },
          body: content,
        }).then(function (r) {
          if (!r.ok) throw new Error('File upload failed: HTTP ' + r.status);
          return tokenUrl;
        });
      });
  }

  // Create the digital article inside the given dossier.
  // Publication/Category are taken from the dossier; Targets are copied from
  // the dossier so the article lands on the same channel/issue.
  function createArticleInDossier(digitalJson, name, feedHeadline, dossier) {
    var ticket = getTicket();
    var pubId = String(dossier.PublicationId || (dossier.Publication && dossier.Publication.Id) || '');
    var catId = String(dossier.CategoryId || (dossier.Category && dossier.Category.Id) || '');
    var dossierId = String(dossier.ID || dossier.Id);

    var dossierTargets = [];
    return callServer('GetObjects', {
      Ticket: ticket, IDs: [dossierId], Lock: false, Rendition: 'none',
      RequestInfo: ['Targets', 'MetaData'], HaveVersions: null, Areas: null, EditionId: null,
    }).then(function (res) {
      var obj = res.Objects && res.Objects[0];
      if (obj) {
        dossierTargets = obj.Targets || [];
        var bm = obj.MetaData && obj.MetaData.BasicMetaData;
        if (bm) {
          pubId = pubId || String((bm.Publication && bm.Publication.Id) || '');
          catId = catId || String((bm.Category && bm.Category.Id) || '');
        }
      }
      return callServer('GetStates', {
        Ticket: ticket, ID: null,
        Publication: { Id: pubId, __classname__: 'Publication' },
        Issue: null,
        Section: catId ? { Id: catId, __classname__: 'Category' } : null,
        Type: 'Article',
      });
    }).then(function (res) {
      var states = (res && res.States) || [];
      if (!states.length) throw new Error('No workflow statuses available for Articles in this Brand/Category');
      var state = states[0];
      return uploadToTransferServer(digitalJson).then(function (fileUrl) {
        return callServer('CreateObjects', {
          Ticket: ticket, Lock: false, Autonaming: true,
          Objects: [{
            __classname__: 'Object',
            MetaData: {
              __classname__: 'MetaData',
              BasicMetaData: {
                __classname__: 'BasicMetaData',
                ID: null, DocumentID: null,
                Name: String(name).slice(0, 60),
                Type: 'Article',
                Publication: { Id: pubId, __classname__: 'Publication' },
                Category: { Id: catId, __classname__: 'Category' },
                ContentSource: null,
              },
              RightsMetaData: null,
              SourceMetaData: null,
              ContentMetaData: {
                __classname__: 'ContentMetaData',
                Format: DIGITAL_MIME,
              },
              WorkflowMetaData: {
                __classname__: 'WorkflowMetaData',
                State: { Id: state.Id, __classname__: 'State' },
              },
              ExtraMetaData: feedHeadline ? [{
                __classname__: 'ExtraMetaData',
                Property: 'C_HEADLINE',
                Values: [feedHeadline],
              }] : [],
            },
            Relations: [{
              __classname__: 'Relation',
              Parent: dossierId, Child: null, Type: 'Contained',
              Placements: null, ParentVersion: null, ChildVersion: null,
              Geometry: null, Rating: null, Targets: null,
            }],
            Pages: null,
            Files: [{
              __classname__: 'Attachment',
              Rendition: 'native',
              Type: DIGITAL_MIME,
              Content: null, FilePath: null,
              FileUrl: fileUrl,
              EditionId: null, ContentSourceFileLink: null, ContentSourceProxyLink: null,
            }],
            Messages: null, Elements: null,
            Targets: dossierTargets,
            Renditions: null, MessageList: null, ObjectLabels: null, Operations: null,
          }],
        });
      });
    });
  }

  // ─── Shared converter UI ───────────────────────────────────────────────────
  var CSS = [
    '.wdab-scroll{position:absolute;top:0;bottom:0;left:0;right:0;overflow-y:auto;-webkit-overflow-scrolling:touch}',
    '.wdab{max-width:640px;margin:0 auto;padding:24px 16px 48px;font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;color:#1e293b}',
    '.wdab-modal .wdab{padding:4px 2px 8px;max-width:none}',
    '.wdab h2{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:#64748b;margin:0 0 14px}',
    '.wdab .wdab-card{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:20px;margin-bottom:16px;box-shadow:0 1px 2px rgba(0,0,0,.04)}',
    '.wdab-modal .wdab .wdab-card{border:0;box-shadow:none;padding:8px 0;margin-bottom:4px}',
    '.wdab label{display:block;font-weight:500;color:#334155;margin:0 0 4px}',
    '.wdab select,.wdab input[type=text]{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;padding:7px 10px;font:inherit;color:#1e293b;background:#fff}',
    '.wdab input[type=file]{width:100%;font:inherit}',
    '.wdab .wdab-row{margin-bottom:14px}',
    '.wdab button.wdab-btn{display:inline-block;border:0;border-radius:8px;padding:9px 16px;font:inherit;font-weight:600;cursor:pointer;background:#2563eb;color:#fff;width:100%}',
    '.wdab button.wdab-btn:disabled{opacity:.4;cursor:not-allowed}',
    '.wdab .wdab-feed input{border-color:#fcd34d;background:#fffbeb}',
    '.wdab .wdab-feed-note{display:inline-block;font-size:11px;color:#b45309;background:#fffbeb;border:1px solid #fde68a;border-radius:4px;padding:1px 6px;margin-left:8px;font-weight:400}',
    '.wdab .wdab-error{color:#dc2626;margin-top:8px;display:none;white-space:pre-wrap}',
    '.wdab .wdab-warn{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;margin-bottom:14px;display:none}',
    '.wdab .wdab-warn ul{margin:6px 0 0;padding-left:18px;color:#92400e}',
    '.wdab .wdab-entries{max-height:220px;overflow-y:auto;border:1px solid #e2e8f0;border-radius:8px;padding:8px 12px;color:#334155}',
    '.wdab .wdab-entries .n{color:#94a3b8;margin-right:6px}',
    '.wdab .wdab-hidden{display:none}',
    '.wdab .wdab-note{color:#64748b;font-size:12px;margin-top:8px}'
  ].join('\n');

  var cssInjected = false;
  function injectCss() {
    if (cssInjected) return;
    var style = document.createElement('style');
    style.textContent = CSS;
    document.head.appendChild(style);
    cssInjected = true;
  }

  function esc(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function formHtml(p, actionLabel) {
    return '<div class="wdab">' +
      '  <div class="wdab-card">' +
      '    <h2>1 — Set up</h2>' +
      '    <div class="wdab-row">' +
      '      <label for="' + p + '-type">Article type</label>' +
      '      <select id="' + p + '-type">' +
      '        <option value="countdown">Type 1 — Numbered countdown (50 → 1)</option>' +
      '        <option value="ascending">Type 2 — Numbered ascending (1 → 50)</option>' +
      '        <option value="crosshead">Type 3 — Crosshead + text (reviews)</option>' +
      '      </select>' +
      '    </div>' +
      '    <div class="wdab-row">' +
      '      <label for="' + p + '-file">Word document (.docx)</label>' +
      '      <input type="file" id="' + p + '-file" accept=".docx">' +
      '    </div>' +
      '    <button class="wdab-btn" id="' + p + '-parse" disabled>Parse Document</button>' +
      '    <p class="wdab-error" id="' + p + '-parse-error"></p>' +
      '  </div>' +
      '  <div id="' + p + '-preview" class="wdab-hidden">' +
      '    <div class="wdab-card">' +
      '      <h2>2 — Detected metadata (editable)</h2>' +
      '      <div class="wdab-row wdab-feed wdab-hidden" id="' + p + '-feed-row">' +
      '        <label>Feed Headline<span class="wdab-feed-note" id="' + p + '-feed-note">Saved to C_HEADLINE</span></label>' +
      '        <input type="text" id="' + p + '-feed">' +
      '      </div>' +
      '      <div class="wdab-row"><label>Article title</label><input type="text" id="' + p + '-title"></div>' +
      '      <div class="wdab-row"><label>Subtitle</label><input type="text" id="' + p + '-subtitle"></div>' +
      '      <div class="wdab-row"><label>Author name</label><input type="text" id="' + p + '-author"></div>' +
      '      <div class="wdab-warn" id="' + p + '-warn"><strong>Unrecognised lines (not included in output):</strong><ul id="' + p + '-warn-list"></ul></div>' +
      '    </div>' +
      '    <div class="wdab-card">' +
      '      <h2>3 — Entries (<span id="' + p + '-count">0</span>)</h2>' +
      '      <div class="wdab-entries" id="' + p + '-entries"></div>' +
      '    </div>' +
      (actionLabel
        ? '<div class="wdab-card"><h2>4 — ' + esc(actionLabel) + '</h2>' +
          '<button class="wdab-btn" id="' + p + '-action">' + esc(actionLabel) + '</button>' +
          '<p class="wdab-error" id="' + p + '-action-error"></p>' +
          '<p class="wdab-note" id="' + p + '-note"></p></div>'
        : '') +
      '  </div>' +
      '</div>';
  }

  // Wires the converter form; returns a controller for reading the result.
  function wireForm(p, onAction) {
    var $ = function (id) { return document.getElementById(p + '-' + id); };
    var state = { parsedData: null, uploadedFilename: '' };

    $('file').addEventListener('change', function () {
      $('parse').disabled = !$('file').files.length;
      $('preview').classList.add('wdab-hidden');
      state.parsedData = null;
      $('parse-error').style.display = 'none';
    });

    $('parse').addEventListener('click', function () {
      var file = $('file').files[0];
      if (!file) return;
      var type = $('type').value;
      $('parse').disabled = true;
      $('parse').textContent = 'Parsing…';
      $('parse-error').style.display = 'none';

      loadMammoth()
        .then(function (mammoth) { return file.arrayBuffer().then(function (buf) { return mammoth.convertToHtml({ arrayBuffer: buf }); }); })
        .then(function (result) {
          var parsed = type === 'crosshead' ? parseCrosshead(result.value) : parseNumbered(result.value);
          state.parsedData = { meta: parsed.meta, entries: parsed.entries, type: type };
          state.uploadedFilename = file.name.replace(/\.docx$/i, '');

          $('feed').value = parsed.meta.feedHeadline;
          $('feed-row').classList.toggle('wdab-hidden', !parsed.meta.feedHeadline);
          $('title').value = parsed.meta.title;
          $('subtitle').value = parsed.meta.subtitle;
          $('author').value = parsed.meta.author;

          if (parsed.meta.unrecognized && parsed.meta.unrecognized.length) {
            $('warn-list').innerHTML = parsed.meta.unrecognized.map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('');
            $('warn').style.display = 'block';
          } else {
            $('warn').style.display = 'none';
          }

          $('count').textContent = parsed.entries.length;
          $('entries').innerHTML = parsed.entries.slice(0, 60).map(function (e, i) {
            return type === 'crosshead'
              ? '<div><span class="n">' + (i + 1) + '.</span>' + esc(e.crosshead || '(no crosshead)') + '</div>'
              : '<div><span class="n">[' + e.number + ']</span>' + esc(e.name) + '</div>';
          }).join('') + (parsed.entries.length > 60 ? '<div class="n">… and ' + (parsed.entries.length - 60) + ' more</div>' : '');

          $('preview').classList.remove('wdab-hidden');
        })
        .catch(function (err) {
          $('parse-error').textContent = 'Error parsing document: ' + err.message;
          $('parse-error').style.display = 'block';
        })
        .then(function () {
          $('parse').disabled = false;
          $('parse').textContent = 'Parse Document';
        });
    });

    if (onAction && $('action')) {
      $('action').addEventListener('click', function () { onAction(controller); });
    }

    var controller = {
      $: $,
      getResult: function () {
        if (!state.parsedData) return null;
        var pm = state.parsedData.meta;
        var pd = pm.deltas || {};
        var meta = {
          feedHeadline: $('feed').value,
          title: $('title').value,
          subtitle: $('subtitle').value,
          author: $('author').value,
          score: pm.score,
          intro: pm.intro,
          introParts: pm.introParts,
          titleDeltas: $('title').value === pm.title ? pd.title : null,
          subtitleDeltas: $('subtitle').value === pm.subtitle ? pd.subtitle : null,
        };
        var template = deepClone(TEMPLATES[state.parsedData.type]);
        var digital = state.parsedData.type === 'crosshead'
          ? buildCrosshead(template, meta, state.parsedData.entries)
          : buildNumbered(template, meta, state.parsedData.entries, state.parsedData.type);
        return { digital: digital, meta: meta, filename: state.uploadedFilename };
      },
    };
    return controller;
  }

  // ─── Apps menu: standalone converter with .digital download ───────────────
  ContentStationSdk.registerCustomApp({
    name: 'word-digital-article-builder',
    title: 'Word → Digital Article',
    content: '<div class="wdab-scroll">' + formHtml('wdab', 'Download .digital file') + '</div>',
    onInit: function () {
      injectCss();
      var form = wireForm('wdab', function (ctl) {
        var result = ctl.getResult();
        if (!result) return;
        var blob = new Blob([JSON.stringify(result.digital)], { type: 'application/json' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = result.filename + '.digital';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      });
      form.$('feed-note').textContent = 'Copy manually to Studio → C_HEADLINE';
    },
  });

  // ─── Dossier toolbar: convert and create directly in this Dossier ─────────
  ContentStationSdk.addDossierToolbarButton({
    label: 'Word → Digital Article',
    onAction: function (config, selection, dossier) {
      injectCss();
      var dialogId = null;
      var busy = false;

      var content = '<div class="wdab-modal">' + formHtml('wdabm', 'Create Digital Article in this Dossier') + '</div>';

      dialogId = ContentStationSdk.openModalDialog({
        title: 'Word → Digital Article',
        subtitle: 'Creates the digital article in Dossier “' + esc(dossier.Name || '') + '”',
        content: content,
        width: 640,
        buttons: [{ label: 'Close', class: 'pale' }],
      });

      wireForm('wdabm', function (ctl) {
        if (busy) return;
        var result = ctl.getResult();
        if (!result) return;
        busy = true;
        var btn = ctl.$('action');
        var errEl = ctl.$('action-error');
        var noteEl = ctl.$('note');
        btn.disabled = true;
        btn.textContent = 'Creating…';
        errEl.style.display = 'none';

        var name = result.meta.title || result.filename;
        createArticleInDossier(JSON.stringify(result.digital), name, result.meta.feedHeadline, dossier)
          .then(function (res) {
            var created = res && res.Objects && res.Objects[0];
            var newName = created && created.MetaData && created.MetaData.BasicMetaData
              ? created.MetaData.BasicMetaData.Name : name;
            ContentStationSdk.showNotification({
              content: 'Digital article “' + esc(newName) + '” created in Dossier “' + esc(dossier.Name || '') + '”.',
              icon: 'check',
            });
            try { ContentStationSdk.refreshCurrentSearch(); } catch (e) { /* non-fatal */ }
            if (dialogId !== null) ContentStationSdk.closeModalDialog(dialogId);
          })
          .catch(function (err) {
            errEl.textContent = err.message + '\nNothing was created. You can fix the issue and try again, or use the Apps-menu version to download the file instead.';
            errEl.style.display = 'block';
          })
          .then(function () {
            busy = false;
            btn.disabled = false;
            btn.textContent = 'Create Digital Article in this Dossier';
          });

        noteEl.textContent = result.meta.feedHeadline ? 'Feed headline will be saved to C_HEADLINE.' : '';
      });
    },
  });
})();
