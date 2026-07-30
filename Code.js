/**
 * IDS PMC Tool — backend
 * Standalone Apps Script project, separate scriptId and separate Google Sheet
 * from the DPR productivity system. Sheet is created by setup() and its ID
 * stored in Script Properties under PMC_SHEET_ID.
 */

// Material Selection (per BOQ line) and Design (per space) are tracked as two
// separate stage lists — see the Phase 4 restructure. STAGES kept as an alias
// of MATERIAL_STAGES for any stray reference; prefer the named constants.
var MATERIAL_STAGES = ['BOQ', 'Selection', 'Order Placement', 'Delivery'];
var DESIGN_STAGES = ['2D', '3D', 'WD'];
var STAGES = MATERIAL_STAGES;

// Execution trades — mirrors the user's real "Legends" tab (Rate List agency column).
var AGENCIES = [
  'Carpentry', 'Flooring', 'Plumbing', 'Painting', 'False Ceiling', 'Fabrication',
  'Windows', 'Electrical', 'Air Conditioning', 'Automation', 'Furnishing',
  'Acoustics', 'Decor', 'Glass work', 'Modular Furniture', 'Sliding profile door'
];

var SHEET_PROP_KEY = 'PMC_SHEET_ID';

// ---------- Materials catalog seed ----------
// Extracted from the user's real "NAGAR HOUSE Interior Estimate.xlsx" (Rate List
// tab, 2026-07-28) — real materials, units and rates, not placeholders. Agency
// per Rate-List category, per the approved plan (editable afterwards in the UI).
var MATERIALS_SEED = [
  { category: 'CARPENTRY', name: '12mm Ply', unit: 'Sq. Ft', rate: 50, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm Ply', unit: 'Sq. Ft', rate: 80, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '12mm HDHMR', unit: 'Sq. Ft', rate: 60, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm HDHMR', unit: 'Sq. Ft', rate: 100, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm Plyboard/ Flush Door', unit: 'Sq. Ft', rate: 120, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '12mm Ply with 4mm hdhmr', unit: 'Sq. Ft', rate: 120, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm Ply with 4mm hdhmr', unit: 'Sq. Ft', rate: 140, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '12mm Ply with 6mm HDHMR (routed)', unit: 'Sq. Ft', rate: 250, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm Ply with 6mm HDHMR (routed)', unit: 'Sq. Ft', rate: 280, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '12mm Flexible Ply', unit: 'Sq. Ft', rate: 150, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '12mm Flexible Ply finish in 4mm hdhmr', unit: 'Sq. Ft', rate: 230, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '12mm Flexible Ply finish in veneer', unit: 'Sq. Ft', rate: 270, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'SF Laminate Finish', unit: 'Sq. Ft', rate: 50, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Mid Range Laminate Finish', unit: 'Sq. Ft', rate: 75, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Top end Laminate Finish', unit: 'Sq. Ft', rate: 100, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm ply + Veneer/ Teak 4mm (Economic)', unit: 'Sq. Ft', rate: 155, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm ply + Veneer/ Teak 4mm (Regular)', unit: 'Sq. Ft', rate: 200, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm ply + Veneer/ Teak 4mm (Decorative - Luxury)', unit: 'Sq. Ft', rate: 280, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '19mm ply + wall paper', unit: 'Sq. Ft', rate: 180, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Veneer/ Teak 4mm (Economic)', unit: 'Sq. Ft', rate: 75, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Veneer/ Teak 4mm (Regular)', unit: 'Sq. Ft', rate: 120, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Veneer/ Teak 4mm (Decorative - Luxury)', unit: 'Sq. Ft', rate: 200, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Decorative Sheet (Economic)', unit: 'Sq. Ft', rate: 120, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Decorative Sheet (Mid Range)', unit: 'Sq. Ft', rate: 180, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Decorative Sheet (Luxury)', unit: 'Sq. Ft', rate: 250, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'WPC / HDHMR mouldings', unit: 'Running Ft', rate: 150, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '45mm Flush Door (30mm + 4mm +4mm) veneer', unit: 'Sq. Ft', rate: 360, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: '45mm Flush Door (30mm + 4mm +4mm) hdhmr', unit: 'Sq. Ft', rate: 140, agency: 'Carpentry' },
  { category: 'CARPENTRY', name: 'Flush Door (25mm + 12mm +12mm MDF) (routed)', unit: 'Sq. Ft', rate: 280, agency: 'Carpentry' },
  { category: 'FLOORING', name: "2'X2' tile", unit: 'Sq. Ft', rate: 65, agency: 'Flooring' },
  { category: 'FLOORING', name: "4'X2' tile", unit: 'Sq. Ft', rate: 100, agency: 'Flooring' },
  { category: 'FLOORING', name: '2\'6"X5\' tile', unit: 'Sq. Ft', rate: 120, agency: 'Flooring' },
  { category: 'FLOORING', name: "4'x6'", unit: 'Sq. Ft', rate: 150, agency: 'Flooring' },
  { category: 'FLOORING', name: "4'x8' tile", unit: 'Sq. Ft', rate: 170, agency: 'Flooring' },
  { category: 'FLOORING', name: '1\'x1\'6" highlighter tile', unit: 'Sq. Ft', rate: 320, agency: 'Flooring' },
  { category: 'FLOORING', name: 'Italian Floor', unit: 'Sq. Ft', rate: 500, agency: 'Flooring' },
  { category: 'FLOORING', name: 'Laminated wooden floor (MDF/HDF base)', unit: 'Sq. Ft', rate: 150, agency: 'Flooring' },
  { category: 'FLOORING', name: 'Engineered wooden floor', unit: 'Sq. Ft', rate: 500, agency: 'Flooring' },
  { category: 'FLOORING', name: 'SPC flooring', unit: 'Sq. Ft', rate: 250, agency: 'Flooring' },
  { category: 'FLOORING', name: 'Carpet Floor', unit: 'Sq. Ft', rate: 120, agency: 'Flooring' },
  { category: 'PAINT & Polish', name: 'Wall pop punning', unit: 'Sq. Ft', rate: 50, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'Wall pop punning (exclusive)', unit: 'Sq. Ft', rate: 150, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'Distemper Paint', unit: 'Sq. Ft', rate: 10, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'Emulsion Paint', unit: 'Sq. Ft', rate: 40, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'Enamel Paint', unit: 'Sq. Ft', rate: 80, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'Texture Paint', unit: 'Sq. Ft', rate: 150, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'Lustre Paint', unit: 'Sq. Ft', rate: 150, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'PU Paint', unit: 'Sq. Ft', rate: 250, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'DUCO Paint', unit: 'Sq. Ft', rate: 150, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'PU Polish', unit: 'Sq. Ft', rate: 250, agency: 'Painting' },
  { category: 'PAINT & Polish', name: 'Melamine Polish', unit: 'Sq. Ft', rate: 150, agency: 'Painting' },
  { category: 'FALSE CEILING', name: 'false ceiling with gyypsum board', unit: 'Sq. Ft', rate: 100, agency: 'False Ceiling' },
  { category: 'FALSE CEILING', name: 'false ceiling with Pop board', unit: 'Sq. Ft', rate: 120, agency: 'False Ceiling' },
  { category: 'FALSE CEILING', name: 'false ceiling with Shera Board', unit: 'Sq. Ft', rate: 80, agency: 'False Ceiling' },
  { category: 'FALSE CEILING', name: 'false ceiling with VOX / Pare', unit: 'Sq. Ft', rate: 250, agency: 'False Ceiling' },
  { category: 'FALSE CEILING', name: 'false ceiling with HPL / ACP', unit: 'Sq. Ft', rate: 300, agency: 'False Ceiling' },
  { category: 'FALSE CEILING', name: 'strech ceiling', unit: 'Sq. Ft', rate: 1500, agency: 'False Ceiling' },
  { category: 'Lighting & Fan', name: 'Wall Lamps', unit: 'L.S.', rate: 5000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Ceiling hanging Lamps', unit: 'L.S.', rate: 5000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Chandelier (Small)', unit: 'L.S.', rate: 15000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Chandelier (Mid Size)', unit: 'L.S.', rate: 25000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Chandelier (Large)', unit: 'L.S.', rate: 75000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Chandelier (Luxury)', unit: 'L.S.', rate: 200000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Fan (economical Range)', unit: 'Nos.', rate: 5000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Fan (Mid Range)', unit: 'Nos.', rate: 10000, agency: 'Electrical' },
  { category: 'Lighting & Fan', name: 'Fan (luxury Range)', unit: 'Nos.', rate: 15000, agency: 'Electrical' },
  { category: 'PARTTIONS', name: 'Internal Glass Partitions', unit: 'Sq. Ft', rate: 2500, agency: 'Fabrication' },
  { category: 'PARTTIONS', name: 'metal work', unit: 'Running Ft', rate: 2500, agency: 'Fabrication' },
  { category: 'PARTTIONS', name: 'Metal Jali', unit: 'Sq. Ft', rate: 3000, agency: 'Fabrication' },
  { category: 'DECOR', name: 'Painting - Artifacts (Economic Range)', unit: 'L.S.', rate: 3000, agency: 'Decor' },
  { category: 'DECOR', name: 'Painting - Artifacts (Mid Range)', unit: 'L.S.', rate: 8000, agency: 'Decor' },
  { category: 'DECOR', name: 'Painting - Artifacts (Luxury Range)', unit: 'L.S.', rate: 15000, agency: 'Decor' },
  { category: 'DECOR', name: 'Wall-paper (Economic Range)', unit: 'Sq. Ft', rate: 60, agency: 'Decor' },
  { category: 'DECOR', name: 'Wall-paper (Mid Range)', unit: 'Sq. Ft', rate: 120, agency: 'Decor' },
  { category: 'DECOR', name: 'Wall-paper (Luxury Range)', unit: 'Sq. Ft', rate: 200, agency: 'Decor' },
  { category: 'DECOR', name: 'Sculpture (Small)', unit: 'L.S.', rate: 5000, agency: 'Decor' },
  { category: 'DECOR', name: 'Sculpture (Mid Size)', unit: 'L.S.', rate: 15000, agency: 'Decor' },
  { category: 'DECOR', name: 'Sculpture (Large)', unit: 'L.S.', rate: 25000, agency: 'Decor' },
  { category: 'DECOR', name: 'planter (small)', unit: 'Nos.', rate: 4000, agency: 'Decor' },
  { category: 'DECOR', name: 'planter (Large)', unit: 'Nos.', rate: 8000, agency: 'Decor' },
  { category: 'DECOR', name: 'Artifacts and table decor', unit: 'L.S.', rate: 4000, agency: 'Decor' },
  { category: 'DECOR', name: 'floor lamp', unit: 'Nos.', rate: 8000, agency: 'Decor' },
  { category: 'DECOR', name: 'Mirror', unit: 'Nos.', rate: 5000, agency: 'Decor' },
  { category: 'DECOR', name: 'Artifacts embellishment', unit: 'L.S.', rate: 10000, agency: 'Decor' },
  { category: 'DECOR', name: 'Fire Place', unit: 'L.S.', rate: 100000, agency: 'Decor' },
  { category: 'CLADDING', name: 'Stone (Economic - cut pieces)', unit: 'Sq. Ft', rate: 150, agency: 'Flooring' },
  { category: 'CLADDING', name: 'Stone (Mid range Decorative Slabs)', unit: 'Sq. Ft', rate: 250, agency: 'Flooring' },
  { category: 'CLADDING', name: 'Stone (Luxury/ Backlit)', unit: 'Sq. Ft', rate: 1500, agency: 'Flooring' },
  { category: 'CLADDING', name: 'HPL Cladding', unit: 'Sq. Ft', rate: 400, agency: 'Flooring' },
  { category: 'CLADDING', name: 'ACP Cladding', unit: 'Sq. Ft', rate: 250, agency: 'Flooring' },
  { category: 'FURNITURE', name: 'TV Console (Runner Only)', unit: 'Sq. Ft', rate: 3000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Wardrobes (Mid Range)', unit: 'Sq. Ft', rate: 3000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Kitchen (Mid Range in HG Laminate)', unit: 'Sq. Ft', rate: 3000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Furniture (Mid Range)', unit: 'Sq. Ft', rate: 3000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Sofa (Luxury)', unit: 'Nos.', rate: 30000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Centre Table', unit: 'L.S.', rate: 30000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Nesting Tables', unit: 'L.S.', rate: 15000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Dining Table (Wooden)', unit: 'Sq. Ft', rate: 3000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Dining Table (Regular Stone)', unit: 'Sq. Ft', rate: 5000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Dining Table (Luxury Stone)', unit: 'Sq. Ft', rate: 10000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Dining Chairs', unit: 'Nos.', rate: 15000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Dining Chairs (luxury)', unit: 'Nos.', rate: 22000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Shelf unit with profile glass shutters (with glass shelves inside)', unit: 'Sq. Ft', rate: 2000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Shelf unit with profile glass shutters (with ply shelves inside)', unit: 'Sq. Ft', rate: 1800, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Console (Economic Range)', unit: 'Nos.', rate: 30000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Console (Mid Range)', unit: 'Nos.', rate: 40000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Console (Luxury Range)', unit: 'Nos.', rate: 50000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Bed (fully hydraulic + full cuahion)', unit: 'Nos.', rate: 100000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'fabric panel', unit: 'Nos.', rate: 450, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: '3 seater sofa', unit: 'Nos.', rate: 75000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: '2 seater sofa', unit: 'Nos.', rate: 50000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'single seater sofa', unit: 'Nos.', rate: 25000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'ootman 2 seater', unit: 'Nos.', rate: 50000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'chair', unit: 'Nos.', rate: 25000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Side Table (Economic Range)', unit: 'Nos.', rate: 8000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Side Table (Mid Range)', unit: 'Nos.', rate: 12000, agency: 'Modular Furniture' },
  { category: 'FURNITURE', name: 'Side Table (Luxury Range)', unit: 'Nos.', rate: 18000, agency: 'Modular Furniture' },
  { category: 'D/W sections', name: 'windows section (economical Range)', unit: 'Sq. Ft', rate: 700, agency: 'Windows' },
  { category: 'D/W sections', name: 'windows section (Mid Range)', unit: 'Sq. Ft', rate: 1200, agency: 'Windows' },
  { category: 'D/W sections', name: 'windows section (luxury Range)', unit: 'Sq. Ft', rate: 1500, agency: 'Windows' },
  { category: 'D/W sections', name: 'Sliding prodile glass (plain)', unit: 'Sq. Ft', rate: 1200, agency: 'Windows' },
  { category: 'D/W sections', name: 'Sliding prodile glass (fabric)', unit: 'Sq. Ft', rate: 1350, agency: 'Windows' },
  { category: 'D/W sections', name: 'Sliding prodile glass (fluted)', unit: 'Sq. Ft', rate: 1500, agency: 'Windows' },
  { category: 'D/W sections', name: 'Clear toughned glass Railing', unit: 'Sq. Ft', rate: 700, agency: 'Windows' },
  { category: 'D/W sections', name: 'Metal door', unit: 'Sq. Ft', rate: 3000, agency: 'Windows' },
  { category: 'Furnishings', name: 'Throws', unit: 'Nos.', rate: 3000, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Throw cushions', unit: 'Nos.', rate: 800, agency: 'Furnishing' },
  { category: 'Furnishings', name: '6" mattress SIZE: 6\'X6\'6"', unit: 'Nos.', rate: 30000, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Rugs (Economic Range)', unit: 'Sq. Ft', rate: 100, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Rugs (Mid Range)', unit: 'Sq. Ft', rate: 250, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Rugs (Luxury Range)', unit: 'Sq. Ft', rate: 750, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Curtains (Economic Range)', unit: 'Sq. Ft', rate: 50, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Curtains (Mid Range)', unit: 'Sq. Ft', rate: 100, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Curtains (Luxury Range)', unit: 'Sq. Ft', rate: 150, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Blinds (Economic Range)', unit: 'Sq. Ft', rate: 120, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Blinds (Mid Range)', unit: 'Sq. Ft', rate: 200, agency: 'Furnishing' },
  { category: 'Furnishings', name: 'Blinds (Luxury Range)', unit: 'Sq. Ft', rate: 250, agency: 'Furnishing' }
];

// Real BOQ line items pulled from the user's Entrance Foyer + Drawing Room
// (same source file) — used to seed a faithful, realistic demo project.
// Length/Width/DepthNos are the exact figures from the user's original
// "BoQ (Quantity Sheet)" takeoff tab (Phase 1 extraction) — L*W*D reproduces
// each line's original Quantity exactly. A few lines carry a demo
// contingency/revisedRate to illustrate those fields (Phase 6).
var DEMO_BOQ_ITEMS = [
  { space: 'Entrance Foyer', name: 'Italian Floor', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 10, width: 8, depthNos: 1, rate: 500 },
  { space: 'Entrance Foyer', name: 'Ply Board Ceiling finish with veneer', category: 'CARPENTRY', agency: 'Carpentry', unit: 'Sq. Ft', length: 10, width: 8, depthNos: 1, rate: 200 },
  { space: 'Entrance Foyer', name: 'Ceiling veneer polish', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 10, width: 8, depthNos: 1, rate: 250 },
  { space: 'Entrance Foyer', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 10, width: 2.5, depthNos: 2, rate: 50 },
  { space: 'Entrance Foyer', name: 'Texture paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 10, width: 2.5, depthNos: 2, rate: 150 },
  { space: 'Entrance Foyer', name: 'Back Lit Stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 3.5, width: 10, depthNos: 1, rate: 500 },
  { space: 'Entrance Foyer', name: 'Venner panel', category: 'CARPENTRY', agency: 'Carpentry', unit: 'Sq. Ft', length: 26, width: 1, depthNos: 1, rate: 200 },
  { space: 'Entrance Foyer', name: 'Venner panel polish', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 26, width: 1, depthNos: 1, rate: 250 },
  { space: 'Entrance Foyer', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 2.5, width: 10, depthNos: 1, rate: 150 },
  { space: 'Entrance Foyer', name: 'Texture paint', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 2.5, width: 10, depthNos: 1, rate: 500 },
  { space: 'Entrance Foyer', name: 'Opening to living room jamming', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 22, width: 1, depthNos: 3, rate: 500 },
  { space: 'Entrance Foyer', name: 'Around Opening to living room stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 1.5, width: 6, depthNos: 1, rate: 500 },
  { space: 'Entrance Foyer', name: 'Around Door to drawing room stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 7, width: 1.5, depthNos: 1, rate: 500 },
  { space: 'Entrance Foyer', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 2, width: 10, depthNos: 2, rate: 50 },
  { space: 'Entrance Foyer', name: 'Texture paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 2, width: 10, depthNos: 2, rate: 40 },
  { space: 'Entrance Foyer', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 2.5, width: 10, depthNos: 1, rate: 40, contingency: 5 },
  { space: 'Entrance Foyer', name: 'Texture paint', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 2.5, width: 10, depthNos: 1, rate: 500 },
  { space: 'Entrance Foyer', name: 'Main Entrance Jamming', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 22, width: 1, depthNos: 3, rate: 500 },
  { space: 'Entrance Foyer', name: 'Around Opening stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 1.5, width: 6, depthNos: 1, rate: 500 },
  { space: 'Entrance Foyer', name: 'Main Entrance Metal door', category: 'D/W sections', agency: 'Windows', unit: 'Sq. Ft', length: 8.5, width: 5, depthNos: 1, rate: 3000, revisedRate: 3200 },
  { space: 'Entrance Foyer', name: 'Console', category: 'FURNITURE', agency: 'Modular Furniture', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 40000 },
  { space: 'Entrance Foyer', name: 'Wall lamp', category: 'DECOR', agency: 'Decor', unit: 'L.S.', length: 1, width: 1, depthNos: 1, rate: 15000 },
  { space: 'Entrance Foyer', name: 'Sitting human figure light', category: 'Lighting & Fan', agency: 'Electrical', unit: 'L.S.', length: 1, width: 1, depthNos: 1, rate: 25000 },
  { space: 'Entrance Foyer', name: 'Chandelier', category: 'DECOR', agency: 'Decor', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 4000 },
  { space: 'Entrance Foyer', name: 'Planter', category: 'DECOR', agency: 'Decor', unit: 'L.S.', length: 1, width: 1, depthNos: 1, rate: 4000 },
  { space: 'Entrance Foyer', name: 'Artifacts and table decor', category: 'DECOR', agency: 'Decor', unit: 'L.S.', length: 1, width: 1, depthNos: 3, rate: 4000 },
  { space: 'Entrance Foyer', name: 'Rug', category: 'Furnishings', agency: 'Furnishing', unit: 'Sq. Ft', length: 4, width: 4, depthNos: 1, rate: 750 },
  { space: 'Drawing Room', name: 'Italian Floor', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 15, width: 18, depthNos: 1, rate: 500, contingency: 15 },
  { space: 'Drawing Room', name: 'Gypsum Board Ceiling', category: 'FALSE CEILING', agency: 'False Ceiling', unit: 'Sq. Ft', length: 15, width: 18, depthNos: 1, rate: 120 },
  { space: 'Drawing Room', name: 'Gypsum Board Ceiling Paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 15, width: 18, depthNos: 1, rate: 80 },
  { space: 'Drawing Room', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 10, width: 8, depthNos: 1, rate: 50 },
  { space: 'Drawing Room', name: 'Texture paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 10, width: 8, depthNos: 1, rate: 40 },
  { space: 'Drawing Room', name: 'Stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 3.5, width: 10, depthNos: 1, rate: 500 },
  { space: 'Drawing Room', name: 'Window to living room stone frame', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 31, width: 1, depthNos: 3, rate: 500 },
  { space: 'Drawing Room', name: 'Metal Jali', category: 'PARTTIONS', agency: 'Fabrication', unit: 'Sq. Ft', length: 8, width: 6, depthNos: 1, rate: 2500 },
  { space: 'Drawing Room', name: 'Window to living room stone panel above window', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 1.5, width: 1, depthNos: 1, rate: 500 },
  { space: 'Drawing Room', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 9, width: 10, depthNos: 1, rate: 50 },
  { space: 'Drawing Room', name: 'Texture paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 9, width: 10, depthNos: 1, rate: 150 },
  { space: 'Drawing Room', name: 'Stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 2, width: 10, depthNos: 1, rate: 500 },
  { space: 'Drawing Room', name: 'Concerte interior sqaure grid and finished in paint', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 1.5, width: 10, depthNos: 1, rate: 320 },
  { space: 'Drawing Room', name: 'Wall punning with grooves', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 16, width: 1.4, depthNos: 1, rate: 50 },
  { space: 'Drawing Room', name: 'Texture paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 16, width: 1.4, depthNos: 1, rate: 150 },
  { space: 'Drawing Room', name: 'Window to balcony stone frame', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 31, width: 1, depthNos: 3, rate: 500 },
  { space: 'Drawing Room', name: 'Around Window to living room stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 3, width: 8, depthNos: 1, rate: 500 },
  { space: 'Drawing Room', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 7.5, width: 10, depthNos: 1, rate: 50 },
  { space: 'Drawing Room', name: 'Texture paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 7.5, width: 10, depthNos: 1, rate: 40 },
  { space: 'Drawing Room', name: 'Stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 2, width: 10, depthNos: 1, rate: 500 },
  { space: 'Drawing Room', name: 'Sliding Glass Window section', category: 'D/W sections', agency: 'Windows', unit: 'Sq. Ft', length: 7, width: 7.5, depthNos: 1, rate: 1200 },
  { space: 'Drawing Room', name: 'Stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 3.5, width: 10, depthNos: 1, rate: 500 },
  { space: 'Drawing Room', name: 'Wall punning', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 4.5, width: 10, depthNos: 1, rate: 50 },
  { space: 'Drawing Room', name: 'Texture paint', category: 'PAINT & Polish', agency: 'Painting', unit: 'Sq. Ft', length: 4.5, width: 10, depthNos: 1, rate: 40 },
  { space: 'Drawing Room', name: 'Around Opening to foyer stone panel', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 6.5, width: 1.5, depthNos: 1, rate: 500 },
  { space: 'Drawing Room', name: 'Sliding Glass Door 12mm sliding toughened glass with powder coated aluminium section with Hardware and handles', category: 'D/W sections', agency: 'Windows', unit: 'Sq. Ft', length: 4, width: 9.5, depthNos: 1, rate: 1200 },
  { space: 'Drawing Room', name: 'Opening stone frame', category: 'FLOORING', agency: 'Flooring', unit: 'Sq. Ft', length: 24, width: 1, depthNos: 3, rate: 500 },
  { space: 'Drawing Room', name: '3 seater sofa', category: 'FURNITURE', agency: 'Modular Furniture', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 75000, revisedRate: 78000 },
  { space: 'Drawing Room', name: '2 seater ottoman', category: 'FURNITURE', agency: 'Modular Furniture', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 50000 },
  { space: 'Drawing Room', name: '3 seater sofa', category: 'FURNITURE', agency: 'Modular Furniture', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 75000 },
  { space: 'Drawing Room', name: 'chair', category: 'FURNITURE', agency: 'Modular Furniture', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 25000 },
  { space: 'Drawing Room', name: 'Center table', category: 'FURNITURE', agency: 'Modular Furniture', unit: 'L.S.', length: 1, width: 1, depthNos: 1, rate: 30000 },
  { space: 'Drawing Room', name: 'Side nesting table', category: 'FURNITURE', agency: 'Modular Furniture', unit: 'L.S.', length: 1, width: 1, depthNos: 1, rate: 15000 },
  { space: 'Drawing Room', name: 'Painting - Wall Hung (3\'6" x 4\'6" in size)', category: 'DECOR', agency: 'Decor', unit: 'L.S.', length: 1, width: 1, depthNos: 1, rate: 15000 },
  { space: 'Drawing Room', name: 'Floor Lamp', category: 'DECOR', agency: 'Decor', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 8000 },
  { space: 'Drawing Room', name: 'Hanging LIght', category: 'Lighting & Fan', agency: 'Electrical', unit: 'L.S.', length: 1, width: 1, depthNos: 1, rate: 5000 },
  { space: 'Drawing Room', name: 'Ceiling Fan', category: 'Lighting & Fan', agency: 'Electrical', unit: 'Nos.', length: 1, width: 1, depthNos: 1, rate: 10000 },
  { space: 'Drawing Room', name: 'Artifacts and table decor', category: 'DECOR', agency: 'Decor', unit: 'L.S.', length: 1, width: 1, depthNos: 10, rate: 4000 },
  { space: 'Drawing Room', name: 'Roman curtain', category: 'Furnishings', agency: 'Furnishing', unit: 'Sq. Ft', length: 7.5, width: 7, depthNos: 1, rate: 150 },
  { space: 'Drawing Room', name: 'Rug', category: 'Furnishings', agency: 'Furnishing', unit: 'Sq. Ft', length: 8, width: 12, depthNos: 1, rate: 250 },
  { space: 'Drawing Room', name: 'Throws', category: 'Furnishings', agency: 'Furnishing', unit: 'Nos.', length: 1, width: 1, depthNos: 3, rate: 3000 },
  { space: 'Drawing Room', name: 'Throw cushions', category: 'Furnishings', agency: 'Furnishing', unit: 'Nos.', length: 1, width: 1, depthNos: 10, rate: 800 },
];

// ---------- Notifications ----------
// Email works out of the box (MailApp, no external account needed). WhatsApp is
// off until WHATSAPP_TOKEN + WHATSAPP_PHONE_ID + WHATSAPP_TO are set as Script
// Properties (Project Settings > Script Properties in the Apps Script editor) —
// sendWhatsApp_ silently no-ops until then, so nothing breaks in the meantime.
var NOTIFY_EMAIL = 'sidinani14@gmail.com';
var APP_URL = 'https://pmc.ideaformdesignstudio.com';

function sendEmail_(subject, body) {
  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (e) {
    Logger.log('sendEmail_ failed: ' + e);
  }
}

function sendWhatsApp_(text) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('WHATSAPP_TOKEN');
  var phoneId = props.getProperty('WHATSAPP_PHONE_ID');
  var to = props.getProperty('WHATSAPP_TO');
  if (!token || !phoneId || !to) return; // not configured yet
  try {
    UrlFetchApp.fetch('https://graph.facebook.com/v20.0/' + phoneId + '/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        messaging_product: 'whatsapp',
        to: to,
        type: 'text',
        text: { body: text }
      })
    });
  } catch (e) {
    Logger.log('sendWhatsApp_ failed: ' + e);
  }
}

function notify_(subject, body) {
  sendEmail_(subject, body);
  sendWhatsApp_(subject + '\n' + body);
}

function projectName_(projectId) {
  var rows = rowsToObjects_(getTab_(getSS_(), 'PROJECTS'));
  var p = rows.filter(function (r) { return r.ProjectID === projectId; })[0];
  return p ? p.Name : projectId;
}
function spaceName_(spaceId) {
  if (!spaceId) return 'General';
  var rows = rowsToObjects_(getTab_(getSS_(), 'SPACES'));
  var s = rows.filter(function (r) { return r.SpaceID === spaceId; })[0];
  return s ? s.Name : spaceId;
}

var TABS = {
  PROJECTS: { name: 'PROJECTS', headers: ['ProjectID', 'Name', 'Address', 'StartDate', 'TargetMoveIn', 'Budget', 'ClientName', 'CreatedAt'] },
  SPACES: { name: 'SPACES', headers: ['SpaceID', 'ProjectID', 'Name', 'SortOrder', 'DesignJSON', 'DesignRollup'] },
  MATERIALS_CONFIG: { name: 'MATERIALS_CONFIG', headers: ['MaterialID', 'Category', 'Name', 'Unit', 'Rate', 'Agency', 'Active', 'CreatedAt'] },
  BOQ_ITEMS: {
    name: 'BOQ_ITEMS', headers: [
      'BoqItemID', 'ProjectID', 'SpaceID', 'MaterialID', 'MaterialName', 'Category', 'Agency',
      'Unit', 'Length', 'Width', 'DepthNos', 'Contingency', 'Rate', 'RevisedRate',
      'StagesJSON', 'RollupStatus', 'Notes', 'CreatedAt', 'UpdatedAt'
    ]
  },
  DAILY_LOG: { name: 'DAILY_LOG', headers: ['LogID', 'Date', 'ProjectID', 'SpaceID', 'Entry', 'LoggedBy', 'HasBlocker', 'CreatedAt', 'UpdatesJSON'] },
  SCHEDULE_ACTIVITIES: {
    name: 'SCHEDULE_ACTIVITIES',
    headers: [
      'ActivityID', 'ProjectID', 'SpaceID', 'Agency', 'StartDate', 'EndDate', 'Status', 'PredecessorsJSON',
      'Notes', 'CreatedAt', 'UpdatedAt', 'PercentComplete', 'OriginalStartDate', 'OriginalEndDate', 'DelayReason'
    ]
  }
};

// ---------- Sheet plumbing ----------

function getSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(SHEET_PROP_KEY);
  if (!id) {
    setup();
    id = props.getProperty(SHEET_PROP_KEY);
  }
  return SpreadsheetApp.openById(id);
}

// Appends a missing header column (and backfills a default for existing rows)
// without disturbing any existing column's position — self-healing so minor
// schema additions never need a one-time manual migration step.
function ensureColumn_(sheet, headerName, defaultValue) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headers.indexOf(headerName) > -1) return;
  var col = lastCol + 1;
  sheet.getRange(1, col).setValue(headerName);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var vals = [];
    for (var i = 0; i < lastRow - 1; i++) vals.push([defaultValue]);
    sheet.getRange(2, col, vals.length, 1).setValues(vals);
  }
}

function getTab_(ss, key) {
  var t = TABS[key];
  var sh = ss.getSheetByName(t.name);
  if (!sh) {
    sh = ss.insertSheet(t.name);
    sh.appendRow(t.headers);
    sh.setFrozenRows(1);
  } else if (key === 'SCHEDULE_ACTIVITIES') {
    ensureColumn_(sh, 'PercentComplete', 0);
    ensureColumn_(sh, 'OriginalStartDate', '');
    ensureColumn_(sh, 'OriginalEndDate', '');
    ensureColumn_(sh, 'DelayReason', '');
  } else if (key === 'DAILY_LOG') {
    ensureColumn_(sh, 'UpdatesJSON', '[]');
  } else if (key === 'SPACES') {
    var defaultDesign = JSON.stringify(DESIGN_STAGES.map(function (s) {
      return { stage: s, status: 'Not Started', target: '', actual: '', note: '' };
    }));
    ensureColumn_(sh, 'DesignJSON', defaultDesign);
    ensureColumn_(sh, 'DesignRollup', 'Not Started');
  }
  return sh;
}

function rowsToObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = row[j];
    obj._row = i + 1;
    out.push(obj);
  }
  return out;
}

function findRowById_(sheet, idCol, id) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var colIdx = headers.indexOf(idCol);
  for (var i = 1; i < values.length; i++) {
    if (values[i][colIdx] === id) return i + 1;
  }
  return -1;
}

function nextId_(prefix, existingIds) {
  var max = 0;
  existingIds.forEach(function (id) {
    var m = /(\d+)$/.exec(String(id));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + '-' + String(max + 1).padStart(3, '0');
}

function computeRollup_(stages) {
  var allDone = stages.every(function (s) { return s.status === 'Done' || s.status === 'N/A'; });
  if (allDone) return 'Done';
  var anyProgress = stages.some(function (s) { return s.status === 'In Progress'; });
  if (anyProgress) return 'In Progress';
  return 'Not Started';
}

function daysAgoISO_(n) {
  var d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
}
function daysFromNowISO_(n) { return daysAgoISO_(-n); }
function dayDiffFromToday_(iso) {
  var today = new Date(Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd') + 'T00:00:00');
  var d = new Date(iso + 'T00:00:00');
  return Math.round((today - d) / 86400000);
}

// ---------- One-time setup ----------

function setup() {
  var props = PropertiesService.getScriptProperties();
  var existing = props.getProperty(SHEET_PROP_KEY);
  var ss;
  if (existing) {
    try { ss = SpreadsheetApp.openById(existing); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('IDS PMC Data');
    props.setProperty(SHEET_PROP_KEY, ss.getId());
    var defaultSheet = ss.getSheets()[0];
    ss.deleteSheet(defaultSheet);
  }
  Object.keys(TABS).forEach(function (key) { getTab_(ss, key); });

  seedMaterialsIfEmpty_(ss);

  var projSheet = getTab_(ss, 'PROJECTS');
  if (rowsToObjects_(projSheet).length === 0) {
    seedDemoProject_(ss);
  }
  Logger.log('PMC sheet ready: ' + ss.getUrl());
  return ss.getUrl();
}

// One-time: run this from the Apps Script editor to move an existing sheet from
// the old fixed-subitem TRACKER/CASHFLOW model to the BOQ_ITEMS model. Safe to
// run more than once (idempotent) — it always leaves a fresh demo project.
// clearContent() on the data range, not deleteRows — deleteRows can throw
// "not possible to delete all non-frozen rows" when the deleted range covers
// every remaining row below a frozen header (an intermittent Sheets API
// restriction), which clearContent never hits since row count is unchanged.
function clearDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
}

function migrateToBoq() {
  var ss = getSS_();
  // BOQ_ITEMS' column layout changed structurally (Phase 6: L/W/D/Contingency/
  // RevisedRate replacing Quantity/Amount/payment-tracking columns) — delete
  // and let getTab_ recreate it with the current header row, rather than just
  // clearing data rows under a stale header.
  ['TRACKER', 'CASHFLOW', 'BOQ_ITEMS'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });
  getTab_(ss, 'BOQ_ITEMS');
  ['PROJECTS', 'SPACES', 'DAILY_LOG', 'MATERIALS_CONFIG', 'SCHEDULE_ACTIVITIES'].forEach(function (key) {
    clearDataRows_(getTab_(ss, key));
  });

  seedMaterialsIfEmpty_(ss);
  seedDemoProject_(ss);
  Logger.log('Migrated to BOQ_ITEMS model: ' + ss.getUrl());
  return ss.getUrl();
}

function seedMaterialsIfEmpty_(ss) {
  var sheet = getTab_(ss, 'MATERIALS_CONFIG');
  if (rowsToObjects_(sheet).length > 0) return;
  var now = new Date();
  var rows = MATERIALS_SEED.map(function (m, i) {
    return ['MAT-' + String(i + 1).padStart(3, '0'), m.category, m.name, m.unit, m.rate, m.agency, true, now];
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function findMaterialByName_(materialsRows, name) {
  return materialsRows.filter(function (m) { return m.Name === name; })[0];
}

function seedDemoProject_(ss) {
  var now = new Date();
  var projectId = 'P-001';
  var projSheet = getTab_(ss, 'PROJECTS');
  projSheet.appendRow([
    projectId, 'Sharma Residence (Demo Project)', 'Kolar Road, Bhopal',
    daysAgoISO_(70), daysFromNowISO_(40), 3200000, 'Rohit & Priya Sharma', now
  ]);

  var spaceSheet = getTab_(ss, 'SPACES');
  var boqSheet = getTab_(ss, 'BOQ_ITEMS');
  var logSheet = getTab_(ss, 'DAILY_LOG');
  var materialsRows = rowsToObjects_(getTab_(ss, 'MATERIALS_CONFIG'));

  var spaceNames = ['Entrance Foyer', 'Drawing Room'];
  var spaceIds = {};
  // Both spaces' design is fully delivered in the demo — the "behind schedule"
  // story lives in procurement/execution, not design, so execution activities
  // (some In Progress/Delayed) stay valid under the WD-gates-execution rule.
  var designDone = JSON.stringify(DESIGN_STAGES.map(function (s, i) {
    var t = daysAgoISO_(90 - i * 10);
    return { stage: s, status: 'Done', target: t, originalTarget: t, actual: daysAgoISO_(85 - i * 10), note: '', delayReason: '' };
  }));
  spaceNames.forEach(function (name, idx) {
    var spaceId = 'S-' + String(idx + 1).padStart(3, '0');
    spaceIds[name] = spaceId;
    spaceSheet.appendRow([spaceId, projectId, name, idx + 1, designDone, 'Done']);
  });

  var counters = {};
  var groups = {}; // '<spaceId>|<agency>' -> { installTargets:[], installStatuses:[] }
  DEMO_BOQ_ITEMS.forEach(function (item) {
    var spaceId = spaceIds[item.space];
    if (!spaceId) return;
    counters[item.space] = (counters[item.space] || 0) + 1;
    var idx = counters[item.space] - 1;
    var mode = item.space === 'Entrance Foyer' ? 'ahead' : 'behind';
    var stages = buildDemoStages_(mode, idx);
    var rollup = computeRollup_(stages);
    var mat = findMaterialByName_(materialsRows, item.name);
    var boqItemId = 'BOQ-' + Utilities.getUuid().slice(0, 6);
    boqSheet.appendRow([
      boqItemId, projectId, spaceId, mat ? mat.MaterialID : '', item.name, item.category, item.agency,
      item.unit, item.length, item.width, item.depthNos, item.contingency || 0, item.rate, item.revisedRate || '',
      JSON.stringify(stages), rollup, '', now, now
    ]);

    var groupKey = spaceId + '|' + item.agency;
    var delivery = stages.filter(function (s) { return s.stage === 'Delivery'; })[0];
    var boqStage = stages.filter(function (s) { return s.stage === 'BOQ'; })[0];
    groups[groupKey] = groups[groupKey] || { spaceId: spaceId, agency: item.agency, installTargets: [], statuses: [], startTargets: [] };
    groups[groupKey].installTargets.push(addDaysISO_(delivery.target, 14));
    groups[groupKey].startTargets.push(boqStage.target);
    groups[groupKey].statuses.push(delivery.status);
  });

  var actSheet = getTab_(ss, 'SCHEDULE_ACTIVITIES');
  var spaceGroupCounters = {};
  Object.keys(groups).forEach(function (key) {
    var g = groups[key];
    // Every demo BOQ line shares the same stage-target dates (buildDemoStages_
    // keys them by stage index only), so groups collapse to identical dates —
    // stagger per space so the chart shows a realistic cascading sequence.
    spaceGroupCounters[g.spaceId] = (spaceGroupCounters[g.spaceId] || 0);
    var offset = spaceGroupCounters[g.spaceId] * 4;
    spaceGroupCounters[g.spaceId]++;
    var originalStart = daysAgoISO_(dayDiffFromToday_(g.startTargets.sort()[0]) - offset);
    var originalEnd = daysAgoISO_(dayDiffFromToday_(g.installTargets.sort().slice(-1)[0]) - offset);
    var status = 'Not Started';
    if (g.statuses.every(function (s) { return s === 'Done'; })) status = 'Done';
    else if (g.statuses.some(function (s) { return s === 'In Progress' || s === 'Done'; })) status = 'In Progress';
    var percentComplete = { 'Done': 100, 'In Progress': 55, 'Not Started': 0 }[status];

    // A couple of In-Progress execution activities in the behind-schedule
    // space get their current end date revised later than original, to show
    // the derived red/yellow states on the Schedule chart too.
    var startDate = originalStart, endDate = originalEnd, delayReason = '';
    if (status === 'In Progress' && g.spaceId !== 'S-001' && offset % 8 === 0) {
      endDate = addDaysISO_(originalEnd, 10);
      delayReason = 'Execution Agency Delay';
    }
    var activityId = 'ACT-' + Utilities.getUuid().slice(0, 6);
    actSheet.appendRow([
      activityId, projectId, g.spaceId, g.agency, startDate, endDate, status, '[]', '', now, now,
      percentComplete, originalStart, originalEnd, delayReason
    ]);
  });

  var engineers = ['Deepak Soni', 'Achal Rathore'];
  var sampleLogs = [
    { d: -1, sp: spaceIds['Entrance Foyer'], by: engineers[0], entry: 'Flooring tiles laid, grouting pending tomorrow. False ceiling frame work started.', blocker: false },
    { d: -1, sp: spaceIds['Drawing Room'], by: engineers[1], entry: 'Wall punning and texture paint in progress. Waiting on stone panel material delivery.', blocker: true },
    { d: -3, sp: spaceIds['Entrance Foyer'], by: engineers[0], entry: 'Veneer panel carpentry work in progress on Wall 1.', blocker: false },
    { d: -5, sp: spaceIds['Drawing Room'], by: engineers[1], entry: 'Metal Jali fabrication delayed due to vendor material shortage.', blocker: true },
    { d: -7, sp: spaceIds['Entrance Foyer'], by: engineers[0], entry: 'Site measurement re-verified for console placement. BOQ finalized with client.', blocker: false }
  ];
  sampleLogs.forEach(function (l) {
    logSheet.appendRow([
      Utilities.getUuid().slice(0, 8), daysAgoISO_(-l.d), projectId, l.sp, l.entry, l.by, l.blocker, new Date(), '[]'
    ]);
  });
}

var DELAY_REASONS_DEMO_ = ['Material Selection Delay', 'Order Placement Delay', 'Material Delivery Delay (Vendor)'];

function buildDemoStages_(mode, idx) {
  // Produce a realistic mixed spread of statuses across the 4 material stages
  // (BOQ, Selection, Order Placement, Delivery) — no literal 'Delayed' status
  // anywhere; delay is purely derived from original-vs-current/today dates.
  // A few rows get their current target deliberately revised later than the
  // original to show the red/yellow derived states, with a delay reason.
  var pattern, revisedIdx;
  if (mode === 'ahead') {
    var patterns = [
      ['Done', 'Done', 'Done', 'Done'],
      ['Done', 'Done', 'Done', 'In Progress'],
      ['Done', 'Done', 'In Progress', 'Not Started'],
      ['Done', 'Done', 'In Progress', 'Not Started'],
      ['Done', 'In Progress', 'Not Started', 'Not Started']
    ];
    pattern = patterns[idx % patterns.length];
    revisedIdx = (idx % patterns.length === 3) ? 2 : -1; // Order Placement revised late
  } else {
    var patterns2 = [
      ['Done', 'In Progress', 'Not Started', 'Not Started'],
      ['Done', 'In Progress', 'Not Started', 'Not Started'],
      ['Done', 'Done', 'In Progress', 'Not Started'],
      ['In Progress', 'Not Started', 'Not Started', 'Not Started'],
      ['Not Started', 'Not Started', 'Not Started', 'Not Started']
    ];
    pattern = patterns2[idx % patterns2.length];
    revisedIdx = (idx % patterns2.length === 1) ? 1 : -1; // Selection revised late
  }
  var stages = [];
  STAGES.forEach(function (stageName, i) {
    var status = pattern[i];
    // Earlier stages' targets sit in the past (due already); later stages'
    // targets sit in the future (not yet due) — a believable mixed spread
    // instead of every open stage automatically reading as overdue.
    var daysAgoForTarget = (mode === 'ahead' ? 30 : 15) - i * 15;
    var originalTarget = daysAgoISO_(daysAgoForTarget);
    var target = originalTarget;
    var note = '';
    var delayReason = '';
    if (i === revisedIdx && status !== 'Done') {
      target = addDaysISO_(originalTarget, 10 + i * 3);
      delayReason = DELAY_REASONS_DEMO_[idx % DELAY_REASONS_DEMO_.length];
      note = 'Revised from original date — see delay reason.';
    }
    var actual = (status === 'Done') ? daysAgoISO_(daysAgoForTarget + 2) : '';
    stages.push({ stage: stageName, status: status, target: target, originalTarget: originalTarget, actual: actual, note: note, delayReason: delayReason });
  });
  return stages;
}

// ---------- Web app entry ----------
// Frontend is hosted separately on GitHub Pages (pmc.ideaformdesignstudio.com) and
// calls this as a plain JSON API — same pattern as the DPR backend. GET for reads
// (?action=...), POST with a JSON body (routed by action) for writes. POST bodies
// must be sent with Content-Type: text/plain from the browser to avoid a CORS
// preflight (Apps Script doesn't answer OPTIONS requests).

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function safeRespond_(fn) {
  try { return respond_({ ok: true, data: fn() }); }
  catch (e) { return respond_({ ok: false, error: String(e && e.message || e) }); }
}

function doGet(e) {
  var p = e && e.parameter ? e.parameter : {};
  if (p.action === 'getAllData') return safeRespond_(getAllData);
  if (p.action === 'getSchema') return safeRespond_(getSchema);
  return respond_({ ok: false, error: 'Unknown action: ' + p.action });
}

function doPost(e) {
  var raw = e && e.postData ? e.postData.contents : '{}';
  var data;
  try { data = JSON.parse(raw || '{}'); } catch (err) { return respond_({ ok: false, error: 'Bad JSON body' }); }
  var routes = {
    addProject: addProject, updateProject: updateProject, addSpace: addSpace,
    addMaterial: addMaterial, updateMaterial: updateMaterial,
    addBoqItem: addBoqItem, updateBoqItem: updateBoqItem,
    updateStage: updateStage, updateDesignStage: updateDesignStage, addDailyLog: addDailyLog,
    updateActivity: updateActivity, setActivityDependencies: setActivityDependencies,
    submitActivityUpdates: submitActivityUpdates
  };
  var fn = routes[data.action];
  if (!fn) return respond_({ ok: false, error: 'Unknown action: ' + data.action });
  return safeRespond_(function () { return fn(data.payload || {}); });
}

// ---------- Client-callable API ----------

function getSchema() {
  return { materialStages: MATERIAL_STAGES, designStages: DESIGN_STAGES, agencies: AGENCIES };
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd');
  return v;
}

function getAllData() {
  var ss = getSS_();
  var projects = rowsToObjects_(getTab_(ss, 'PROJECTS')).map(function (r) {
    r.StartDate = normDate_(r.StartDate);
    r.TargetMoveIn = normDate_(r.TargetMoveIn);
    return r;
  });
  var spaces = rowsToObjects_(getTab_(ss, 'SPACES')).map(function (r) {
    r.Design = JSON.parse(r.DesignJSON || '[]').map(function (s) {
      s.target = normDate_(s.target);
      s.actual = normDate_(s.actual);
      s.originalTarget = normDate_(s.originalTarget) || s.target;
      return s;
    });
    return r;
  });
  var materialsConfig = rowsToObjects_(getTab_(ss, 'MATERIALS_CONFIG'));
  var boqItems = rowsToObjects_(getTab_(ss, 'BOQ_ITEMS')).map(function (r) {
    r.Stages = JSON.parse(r.StagesJSON || '[]').map(function (s) {
      s.target = normDate_(s.target);
      s.actual = normDate_(s.actual);
      s.originalTarget = normDate_(s.originalTarget) || s.target;
      return s;
    });
    r.Quantity = (Number(r.Length) || 0) * (Number(r.Width) || 0) * (Number(r.DepthNos) || 0) + (Number(r.Contingency) || 0);
    r.EffectiveRate = r.RevisedRate !== '' && r.RevisedRate !== undefined && r.RevisedRate !== null ? Number(r.RevisedRate) : Number(r.Rate) || 0;
    r.Amount = r.Quantity * r.EffectiveRate;
    return r;
  });
  var dailyLog = rowsToObjects_(getTab_(ss, 'DAILY_LOG')).map(function (r) {
    r.Date = normDate_(r.Date);
    return r;
  });
  dailyLog.sort(function (a, b) { return new Date(b.Date) - new Date(a.Date) || new Date(b.CreatedAt) - new Date(a.CreatedAt); });

  var activities = rowsToObjects_(getTab_(ss, 'SCHEDULE_ACTIVITIES')).map(function (r) {
    r.StartDate = normDate_(r.StartDate);
    r.EndDate = normDate_(r.EndDate);
    r.OriginalStartDate = normDate_(r.OriginalStartDate) || r.StartDate;
    r.OriginalEndDate = normDate_(r.OriginalEndDate) || r.EndDate;
    r.Predecessors = JSON.parse(r.PredecessorsJSON || '[]');
    return r;
  });

  var categories = (function () {
    var seen = {}, out = [];
    materialsConfig.forEach(function (m) { if (!seen[m.Category]) { seen[m.Category] = true; out.push(m.Category); } });
    return out;
  })();

  return {
    schema: { materialStages: MATERIAL_STAGES, designStages: DESIGN_STAGES, agencies: AGENCIES, categories: categories },
    projects: projects,
    spaces: spaces,
    materialsConfig: materialsConfig,
    boqItems: boqItems,
    dailyLog: dailyLog,
    activities: activities
  };
}

function addProject(payload) {
  var ss = getSS_();
  var projSheet = getTab_(ss, 'PROJECTS');
  var existingIds = rowsToObjects_(projSheet).map(function (r) { return r.ProjectID; });
  var projectId = nextId_('P', existingIds);
  projSheet.appendRow([
    projectId, payload.name, payload.address || '', payload.startDate || '',
    payload.targetMoveIn || '', Number(payload.budget) || 0, payload.clientName || '', new Date()
  ]);

  var spaceSheet = getTab_(ss, 'SPACES');
  var spaceNames = (payload.spaces && payload.spaces.length) ? payload.spaces : ['Drawing Room'];
  spaceNames.forEach(function (name, idx) {
    addSpaceInternal_(spaceSheet, projectId, name, idx + 1);
  });
  return { projectId: projectId };
}

function updateProject(payload) {
  var ss = getSS_();
  var sheet = getTab_(ss, 'PROJECTS');
  var row = findRowById_(sheet, 'ProjectID', payload.projectId);
  if (row < 0) throw new Error('Project not found');
  var headers = sheet.getDataRange().getValues()[0];
  var current = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  var map = { Name: 'name', Address: 'address', StartDate: 'startDate', TargetMoveIn: 'targetMoveIn', Budget: 'budget', ClientName: 'clientName' };
  headers.forEach(function (h, i) {
    if (map[h] && payload[map[h]] !== undefined && payload[map[h]] !== '') {
      current[i] = (h === 'Budget') ? Number(payload[map[h]]) : payload[map[h]];
    }
  });
  sheet.getRange(row, 1, 1, headers.length).setValues([current]);
  return { ok: true };
}

function addSpaceInternal_(spaceSheet, projectId, name, sortOrder) {
  var existingIds = rowsToObjects_(spaceSheet).map(function (r) { return r.SpaceID; });
  var spaceId = nextId_('S', existingIds);
  var design = JSON.stringify(DESIGN_STAGES.map(function (s) {
    return { stage: s, status: 'Not Started', target: '', actual: '', note: '' };
  }));
  spaceSheet.appendRow([spaceId, projectId, name, sortOrder, design, 'Not Started']);
  return spaceId;
}

function addSpace(payload) {
  var ss = getSS_();
  var spaceSheet = getTab_(ss, 'SPACES');
  var existing = rowsToObjects_(spaceSheet).filter(function (r) { return r.ProjectID === payload.projectId; });
  var sortOrder = existing.length + 1;
  var spaceId = addSpaceInternal_(spaceSheet, payload.projectId, payload.name, sortOrder);
  return { spaceId: spaceId };
}

// ---------- Materials catalog CRUD ----------

function addMaterial(payload) {
  // payload: { category, name, unit, rate, agency }
  var ss = getSS_();
  var sheet = getTab_(ss, 'MATERIALS_CONFIG');
  var existingIds = rowsToObjects_(sheet).map(function (r) { return r.MaterialID; });
  var materialId = nextId_('MAT', existingIds);
  sheet.appendRow([
    materialId, payload.category, payload.name, payload.unit,
    Number(payload.rate) || 0, payload.agency, true, new Date()
  ]);
  return { materialId: materialId };
}

function updateMaterial(payload) {
  // payload: { materialId, category, name, unit, rate, agency, active }
  var ss = getSS_();
  var sheet = getTab_(ss, 'MATERIALS_CONFIG');
  var row = findRowById_(sheet, 'MaterialID', payload.materialId);
  if (row < 0) throw new Error('Material not found');
  var headers = sheet.getDataRange().getValues()[0];
  var current = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  var map = { Category: 'category', Name: 'name', Unit: 'unit', Rate: 'rate', Agency: 'agency', Active: 'active' };
  headers.forEach(function (h, i) {
    if (map[h] && payload[map[h]] !== undefined) {
      current[i] = (h === 'Rate') ? Number(payload[map[h]]) : payload[map[h]];
    }
  });
  sheet.getRange(row, 1, 1, headers.length).setValues([current]);
  return { ok: true };
}

// ---------- BOQ line items ----------

function addBoqItem(payload) {
  // payload: { projectId, spaceId, materialId, length, width, depthNos, contingency, revisedRate }
  var ss = getSS_();
  var matSheet = getTab_(ss, 'MATERIALS_CONFIG');
  var material = rowsToObjects_(matSheet).filter(function (m) { return m.MaterialID === payload.materialId; })[0];
  if (!material) throw new Error('Material not found');

  var length = Number(payload.length) || 0;
  var width = Number(payload.width) || 0;
  var depthNos = Number(payload.depthNos) || 0;
  var contingency = Number(payload.contingency) || 0;
  var rate = Number(material.Rate) || 0;
  var revisedRate = payload.revisedRate !== undefined && payload.revisedRate !== '' ? Number(payload.revisedRate) : '';
  var stages = STAGES.map(function (stageName) {
    return { stage: stageName, status: 'Not Started', target: '', actual: '', note: '' };
  });

  var boqSheet = getTab_(ss, 'BOQ_ITEMS');
  var boqItemId = 'BOQ-' + Utilities.getUuid().slice(0, 6);
  var now = new Date();
  boqSheet.appendRow([
    boqItemId, payload.projectId, payload.spaceId, material.MaterialID, material.Name, material.Category, material.Agency,
    material.Unit, length, width, depthNos, contingency, rate, revisedRate,
    JSON.stringify(stages), 'Not Started', '', now, now
  ]);
  ensureActivityForBoqItem_(ss, payload.projectId, payload.spaceId, material.Agency);
  return { boqItemId: boqItemId };
}

// ---------- Schedule / PERT activities ----------

function ensureActivityForBoqItem_(ss, projectId, spaceId, agency) {
  var sheet = getTab_(ss, 'SCHEDULE_ACTIVITIES');
  var all = rowsToObjects_(sheet);
  var existing = all.filter(function (a) {
    return a.ProjectID === projectId && a.SpaceID === spaceId && a.Agency === agency;
  })[0];
  if (existing) return existing.ActivityID;

  // Chain after the most-recently-created activity already in this space, if
  // any — rooms run one trade after another by default, not in parallel.
  var spaceActivities = all.filter(function (a) { return a.ProjectID === projectId && a.SpaceID === spaceId; })
    .sort(function (a, b) { return new Date(b.CreatedAt) - new Date(a.CreatedAt); });
  var predecessor = spaceActivities[0];

  var activityId = 'ACT-' + Utilities.getUuid().slice(0, 6);
  var now = new Date();
  // normDate_ first: a date read back via rowsToObjects_ may be a raw Sheets
  // Date object (internally IST-midnight-based, since the spreadsheet's
  // timeZone is Asia/Kolkata) rather than a plain string — arithmetic on the
  // unnormalized value silently lands a day off.
  var startDate = predecessor ? normDate_(predecessor.EndDate) : daysFromNowISO_(0);
  var endDate = addDaysISO_(startDate, 7);
  var predecessors = predecessor ? [predecessor.ActivityID] : [];
  sheet.appendRow([
    activityId, projectId, spaceId, agency, startDate, endDate,
    'Not Started', JSON.stringify(predecessors), '', now, now, 0, startDate, endDate, ''
  ]);
  return activityId;
}

function addDaysISO_(iso, n) {
  var normalized = normDate_(iso) || iso;
  var d = new Date(normalized);
  d.setUTCDate(d.getUTCDate() + n);
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

// Shared write path for a single activity's status/dates/notes/%, used by both
// updateActivity (one activity at a time) and submitActivityUpdates (a batch
// from the daily-log form) — keeps the Done-flip and cascade logic in one place.
function getSpaceDesignStage_(ss, spaceId, stageName) {
  var space = rowsToObjects_(getTab_(ss, 'SPACES')).filter(function (s) { return s.SpaceID === spaceId; })[0];
  if (!space) return null;
  var design = JSON.parse(space.DesignJSON || '[]');
  return design.filter(function (d) { return d.stage === stageName; })[0] || null;
}

function applyActivityChange_(ss, activityId, changes) {
  var sheet = getTab_(ss, 'SCHEDULE_ACTIVITIES');
  var row = findRowById_(sheet, 'ActivityID', activityId);
  if (row < 0) throw new Error('Activity not found: ' + activityId);
  var headers = sheet.getDataRange().getValues()[0];
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });
  var current = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  // Execution can be scheduled/dragged freely, but can't actually start until
  // the space's final design (WD) is delivered.
  if (changes.status === 'In Progress' || changes.status === 'Done') {
    var wd = getSpaceDesignStage_(ss, current[idx.SpaceID], 'WD');
    if (!wd || wd.status !== 'Done') {
      throw new Error('Design (WD) must be delivered for this space before execution can start.');
    }
  }

  if (changes.startDate !== undefined) current[idx.StartDate] = changes.startDate;
  if (changes.endDate !== undefined) current[idx.EndDate] = changes.endDate;
  if (changes.status !== undefined) current[idx.Status] = changes.status;
  if (changes.notes !== undefined) current[idx.Notes] = changes.notes;
  if (changes.percentComplete !== undefined) current[idx.PercentComplete] = Number(changes.percentComplete);
  if (changes.delayReason !== undefined) current[idx.DelayReason] = changes.delayReason;
  current[idx.UpdatedAt] = new Date();
  if (changes.status === 'Done') current[idx.PercentComplete] = 100;
  sheet.getRange(row, 1, 1, headers.length).setValues([current]);

  cascadeFromActivity_(ss, activityId);
}

function updateActivity(payload) {
  // payload: { activityId, startDate?, endDate?, status?, notes? }
  var ss = getSS_();
  applyActivityChange_(ss, payload.activityId, payload);
  return { ok: true };
}

function submitActivityUpdates(payload) {
  // payload: { date, projectId, spaceId, loggedBy, hasBlocker, notes, updates: [{activityId, status, endDate, percentComplete}] }
  var ss = getSS_();
  var summaryParts = [];
  (payload.updates || []).forEach(function (u) {
    applyActivityChange_(ss, u.activityId, { status: u.status, endDate: u.endDate, percentComplete: u.percentComplete });
    var actRow = rowsToObjects_(getTab_(ss, 'SCHEDULE_ACTIVITIES')).filter(function (a) { return a.ActivityID === u.activityId; })[0];
    summaryParts.push((actRow ? actRow.Agency : u.activityId) + ': ' + u.status + ' (' + u.percentComplete + '%)');
  });
  var entry = summaryParts.join('; ') + (payload.notes ? ' — ' + payload.notes : '');

  var logSheet = getTab_(ss, 'DAILY_LOG');
  var logId = 'LOG-' + Utilities.getUuid().slice(0, 8);
  logSheet.appendRow([
    logId, payload.date, payload.projectId, payload.spaceId || '',
    entry, payload.loggedBy, !!payload.hasBlocker, new Date(), JSON.stringify(payload.updates || [])
  ]);

  var proj = projectName_(payload.projectId);
  var space = spaceName_(payload.spaceId);
  var subject = (payload.hasBlocker ? '⚠️ BLOCKER — ' : 'Daily log — ') + proj + ' (' + space + ')';
  var body = payload.loggedBy + ' logged an update on ' + payload.date + ':\n\n' + entry + '\n\n' + APP_URL;
  notify_(subject, body);

  return { logId: logId };
}

function setActivityDependencies(payload) {
  // payload: { activityId, predecessors: [ActivityID, ...] }
  var ss = getSS_();
  var sheet = getTab_(ss, 'SCHEDULE_ACTIVITIES');
  var row = findRowById_(sheet, 'ActivityID', payload.activityId);
  if (row < 0) throw new Error('Activity not found');
  var headers = sheet.getDataRange().getValues()[0];
  var predColIdx = headers.indexOf('PredecessorsJSON');
  var updatedColIdx = headers.indexOf('UpdatedAt');
  var predecessors = (payload.predecessors || []).filter(function (id) { return id !== payload.activityId; });
  sheet.getRange(row, predColIdx + 1).setValue(JSON.stringify(predecessors));
  sheet.getRange(row, updatedColIdx + 1).setValue(new Date());

  predecessors.forEach(function (predId) { cascadeFromActivity_(ss, predId); });
  return { ok: true };
}


// Finish-to-start cascade: any activity that lists `activityId` as a predecessor
// must start no earlier than that predecessor's end. If violated, shift the
// dependent forward (preserving its own duration) and recurse into its own
// dependents. `visited` guards against an accidental cycle.
function cascadeFromActivity_(ss, activityId, visited) {
  visited = visited || {};
  if (visited[activityId]) return;
  visited[activityId] = true;

  var sheet = getTab_(ss, 'SCHEDULE_ACTIVITIES');
  var all = rowsToObjects_(sheet);
  var predecessor = all.filter(function (a) { return a.ActivityID === activityId; })[0];
  if (!predecessor) return;
  var predEnd = normDate_(predecessor.EndDate) || predecessor.EndDate;

  var headers = sheet.getDataRange().getValues()[0];
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  all.forEach(function (successor) {
    var preds = JSON.parse(successor.PredecessorsJSON || '[]');
    if (preds.indexOf(activityId) === -1) return;
    var succStart = normDate_(successor.StartDate) || successor.StartDate;
    var succEnd = normDate_(successor.EndDate) || successor.EndDate;
    if (succStart >= predEnd) return; // already satisfied

    var deltaMs = new Date(predEnd) - new Date(succStart);
    var newStart = shiftDateISO_(succStart, deltaMs);
    var newEnd = shiftDateISO_(succEnd, deltaMs);

    var current = sheet.getRange(successor._row, 1, 1, headers.length).getValues()[0];
    current[idx.StartDate] = newStart;
    current[idx.EndDate] = newEnd;
    current[idx.UpdatedAt] = new Date();
    sheet.getRange(successor._row, 1, 1, headers.length).setValues([current]);

    cascadeFromActivity_(ss, successor.ActivityID, visited);
  });
}

function shiftDateISO_(iso, deltaMs) {
  var d = new Date(iso);
  d.setTime(d.getTime() + deltaMs);
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
}

function updateBoqItem(payload) {
  // payload: { boqItemId, length?, width?, depthNos?, contingency?, revisedRate?, notes? }
  var ss = getSS_();
  var sheet = getTab_(ss, 'BOQ_ITEMS');
  var row = findRowById_(sheet, 'BoqItemID', payload.boqItemId);
  if (row < 0) throw new Error('BOQ item not found');
  var headers = sheet.getDataRange().getValues()[0];
  var current = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });

  if (payload.length !== undefined) current[idx.Length] = Number(payload.length);
  if (payload.width !== undefined) current[idx.Width] = Number(payload.width);
  if (payload.depthNos !== undefined) current[idx.DepthNos] = Number(payload.depthNos);
  if (payload.contingency !== undefined) current[idx.Contingency] = Number(payload.contingency);
  if (payload.revisedRate !== undefined) current[idx.RevisedRate] = payload.revisedRate === '' ? '' : Number(payload.revisedRate);
  if (payload.notes !== undefined) current[idx.Notes] = payload.notes;
  current[idx.UpdatedAt] = new Date();
  sheet.getRange(row, 1, 1, headers.length).setValues([current]);
  return { ok: true };
}

// Shared write path for a single stage object (BOQ or Design) — captures
// `originalTarget` the first time a target is ever saved (never overwritten
// after), and carries an optional delay reason.
function applyStageChange_(stageObj, payload) {
  if (payload.status !== undefined) stageObj.status = payload.status;
  if (payload.target !== undefined) {
    if (!stageObj.originalTarget) stageObj.originalTarget = payload.target;
    stageObj.target = payload.target;
  }
  if (payload.actual !== undefined) stageObj.actual = payload.actual;
  if (payload.note !== undefined) stageObj.note = payload.note;
  if (payload.delayReason !== undefined) stageObj.delayReason = payload.delayReason;
}

function updateStage(payload) {
  // payload: { boqItemId, stage, status, target, actual, note, delayReason }
  var ss = getSS_();
  var sheet = getTab_(ss, 'BOQ_ITEMS');
  var row = findRowById_(sheet, 'BoqItemID', payload.boqItemId);
  if (row < 0) throw new Error('BOQ item not found');
  var headers = sheet.getDataRange().getValues()[0];
  var stagesColIdx = headers.indexOf('StagesJSON');
  var rollupColIdx = headers.indexOf('RollupStatus');
  var updatedColIdx = headers.indexOf('UpdatedAt');
  var stagesRaw = sheet.getRange(row, stagesColIdx + 1).getValue();
  var stages = JSON.parse(stagesRaw || '[]');
  var stageObj = stages.filter(function (s) { return s.stage === payload.stage; })[0];
  if (!stageObj) throw new Error('Stage not found');
  applyStageChange_(stageObj, payload);

  var rollup = computeRollup_(stages);
  sheet.getRange(row, stagesColIdx + 1).setValue(JSON.stringify(stages));
  sheet.getRange(row, rollupColIdx + 1).setValue(rollup);
  sheet.getRange(row, updatedColIdx + 1).setValue(new Date());
  return { rollup: rollup, stages: stages };
}

function updateDesignStage(payload) {
  // payload: { spaceId, stage, status, target, actual, note, delayReason }
  var ss = getSS_();
  var sheet = getTab_(ss, 'SPACES');
  var row = findRowById_(sheet, 'SpaceID', payload.spaceId);
  if (row < 0) throw new Error('Space not found');
  var headers = sheet.getDataRange().getValues()[0];
  var designColIdx = headers.indexOf('DesignJSON');
  var rollupColIdx = headers.indexOf('DesignRollup');
  var stages = JSON.parse(sheet.getRange(row, designColIdx + 1).getValue() || '[]');
  var stageObj = stages.filter(function (s) { return s.stage === payload.stage; })[0];
  if (!stageObj) throw new Error('Design stage not found');
  applyStageChange_(stageObj, payload);

  var rollup = computeRollup_(stages);
  sheet.getRange(row, designColIdx + 1).setValue(JSON.stringify(stages));
  sheet.getRange(row, rollupColIdx + 1).setValue(rollup);

  if (payload.stage === 'WD' && stageObj.status === 'Done') {
    var wdDoneDate = normDate_(stageObj.actual) || stageObj.actual || todayISO_();
    gateExecutionOnDesign_(ss, payload.spaceId, wdDoneDate);
  }
  return { rollup: rollup, stages: stages };
}

function todayISO_() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
}

// Once a space's final design is delivered, any execution activity still
// scheduled to start before that date auto-shifts to start right at it
// (preserving its own duration), then cascades into its own dependents via
// the existing predecessor-chain logic.
function gateExecutionOnDesign_(ss, spaceId, wdDoneDate) {
  var sheet = getTab_(ss, 'SCHEDULE_ACTIVITIES');
  var headers = sheet.getDataRange().getValues()[0];
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });
  var acts = rowsToObjects_(sheet).filter(function (a) { return a.SpaceID === spaceId; });

  acts.forEach(function (a) {
    var start = normDate_(a.StartDate) || a.StartDate;
    if (start >= wdDoneDate) return;
    var end = normDate_(a.EndDate) || a.EndDate;
    var deltaMs = new Date(wdDoneDate) - new Date(start);
    var newStart = shiftDateISO_(start, deltaMs);
    var newEnd = shiftDateISO_(end, deltaMs);

    var current = sheet.getRange(a._row, 1, 1, headers.length).getValues()[0];
    current[idx.StartDate] = newStart;
    current[idx.EndDate] = newEnd;
    current[idx.UpdatedAt] = new Date();
    sheet.getRange(a._row, 1, 1, headers.length).setValues([current]);

    cascadeFromActivity_(ss, a.ActivityID);
  });
}

// ---------- Daily log ----------

function addDailyLog(payload) {
  // payload: { date, projectId, spaceId, entry, loggedBy, hasBlocker }
  var ss = getSS_();
  var sheet = getTab_(ss, 'DAILY_LOG');
  var logId = 'LOG-' + Utilities.getUuid().slice(0, 8);
  sheet.appendRow([
    logId, payload.date, payload.projectId, payload.spaceId || '',
    payload.entry, payload.loggedBy, !!payload.hasBlocker, new Date(), '[]'
  ]);

  var proj = projectName_(payload.projectId);
  var space = spaceName_(payload.spaceId);
  var subject = (payload.hasBlocker ? '⚠️ BLOCKER — ' : 'Daily log — ') + proj + ' (' + space + ')';
  var body = payload.loggedBy + ' logged an update on ' + payload.date + ':\n\n' +
    payload.entry + '\n\n' + APP_URL;
  notify_(subject, body);

  return { logId: logId };
}

// ---------- Overdue digest (run daily via a time trigger) ----------

function checkOverdueAndNotify() {
  var data = getAllData();
  var today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  var lines = [];
  data.boqItems.forEach(function (r) {
    var sp = data.spaces.filter(function (s) { return s.SpaceID === r.SpaceID; })[0];
    var proj = data.projects.filter(function (p) { return p.ProjectID === r.ProjectID; })[0];
    r.Stages.forEach(function (s) {
      var isOverdue = s.target && s.target < today && s.status !== 'Done' && s.status !== 'N/A';
      if (s.status === 'Delayed' || isOverdue) {
        lines.push((proj ? proj.Name : '') + ' — ' + (sp ? sp.Name : '') + ' — ' +
          r.Category + ' / ' + r.MaterialName + ' (' + s.stage + '): ' +
          (s.status === 'Delayed' ? 'marked Delayed' : 'overdue, target was ' + s.target));
      }
    });
  });
  if (!lines.length) return;
  var subject = lines.length + ' delayed/overdue item' + (lines.length === 1 ? '' : 's') + ' across all projects';
  var body = lines.join('\n') + '\n\n' + APP_URL;
  notify_(subject, body);
}

// One-time: run this once from the Apps Script editor to schedule the daily digest.
function setupDailyOverdueTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkOverdueAndNotify') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkOverdueAndNotify').timeBased().everyDays(1).atHour(9).create();
}
