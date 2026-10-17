import json
D = json.load(open('/home/claude/data_v1.json'))

# TRL per item. Judged against use FROM SPACE for the stated application, on the NASA 1–9 scale.
# Tuple: (trl, note or None, verify flag). The note explains any non-obvious call.
T = {
 'T01':(9,None,False),'T02':(9,None,False),'T03':(9,None,False),'T04':(9,None,False),
 'T05':(9,'OCO-2 and TROPOMI retrievals are operational; FLEX, the dedicated mission, is not yet launched.',False),
 'T06':(9,None,False),'T07':(9,None,False),'T08':(9,None,False),'T09':(9,None,False),'T10':(9,None,False),
 'T11':(9,None,False),'T12':(9,None,False),'T13':(9,None,False),'T14':(9,None,False),'T15':(9,None,False),
 'T16':(9,None,False),'T17':(9,None,False),'T18':(9,None,False),'T19':(9,None,False),'T20':(9,None,False),
 'T21':(9,None,False),'T22':(9,None,False),'T23':(9,None,False),'T24':(9,None,False),
 'T25':(8,'Flown and modeled in operational sounders, but as a correction; the field itself is not a delivered product.',False),
 'T26':(7,'Broadband longwave is TRL 9 (CERES). Spectrally resolved far-IR is TRL 7: PREFIRE CubeSats flew 2024; FORUM not yet.',True),
 'T27':(9,None,False),'T28':(9,None,False),'T29':(9,None,False),'T30':(9,None,False),'T31':(9,None,False),
 'T32':(9,None,False),'T33':(9,None,False),
 'T34':(6,'Flight data exists (GLM, ISS-LIS, ASIM) but the height retrieval is research-grade, not an operational product.',False),
 'T35':(9,None,False),'T36':(9,None,False),
 'T37':(7,'Detected routinely after events in research; no operational warning product.',False),
 'T38':(9,None,False),'T39':(9,None,False),'T40':(9,None,False),
 'T41':(8,'Experimental modes on TerraSAR-X and Sentinel-1; no operational along-track product.',False),
 'T42':(9,None,False),
 'T43':(7,'Demonstrated with PALSAR data as a calibration byproduct.',False),
 'T44':(7,'TanDEM-X experiments; BIOMASS tomographic mode in commissioning.',True),
 'T45':(8,'L-band is TRL 9. P-band from orbit is first flown on BIOMASS (2025); 8 until data are proven.',True),
 'T46':(9,None,False),'T47':(9,None,False),'T48':(9,None,False),'T49':(9,None,False),'T50':(9,None,False),
 'T51':(9,'Mission-proven at Mars and the Moon; not applicable to Earth from orbit.',False),
 'T52':(9,None,False),'T53':(9,None,False),'T54':(9,None,False),
 'T55':(9,'Single mission (Aeolus, 2018–2023); laser degradation was a known issue; Aeolus-2 in development.',False),
 'T56':(8,'ACDL on DQ-1 (China, 2022) is reported flown with published data; I have less visibility into performance.',True),
 'T57':(9,None,False),'T58':(9,None,False),'T59':(9,None,False),'T60':(9,None,False),'T61':(9,None,False),
 'T62':(9,None,False),'T63':(9,None,False),'T64':(9,None,False),
 'T65':(6,'Mini-EUSO and TUS operated in orbit and saw TLEs and meteors; a confirmed UHECR detection from orbit is not established.',True),
 'T66':(9,None,False),'T67':(9,None,False),
 'T68':(7,'SEXTANT demonstrated on the ISS (2018).',False),
 'T69':(9,None,False),'T70':(9,None,False),
 'T71':(5,'Airborne campaigns only.',False),
 'T72':(6,'TDS-1 altimetry demonstrated in orbit at research grade; ground GNSS-IR snow depth is operational.',False),
 'T73':(9,None,False),
 'T74':(5,'Ground and airborne demos, including Starlink-illuminated passive SAR; no orbital receiver.',True),
 'T75':(5,'Characterized on OICETS, LCRD and EDRS links; not yet used as a sensing product.',False),
 'T76':(5,'GeoSTAR ground prototype; GIMS (China) status uncertain.',True),
 'T77':(6,'Cold-atom systems operated in orbit (CAL, CACES); accelerometer-grade performance for gravimetry not yet shown.',True),
 'T78':(7,'On-orbit star-tracker data exploited for SDA.',False),
 'T79':(3,'Ground Raman lidar is TRL 9. From orbit the photon budget fails by ~10³; studies only.',False),
 'T80':(5,'Airborne oceanographic and oil-spill fluorosensors.',False),
}
C = {
 'C01':(4,'Lab and field demos of Rydberg receivers; no space-environment validation.',False),
 'C02':(4,'Airborne LDV validates the sensor, not the orbital speckle-decorrelation problem.',False),
 'C03':(3,'Ground LGS magnetometry proves the physics; space geometry untested.',False),
 'C04':(2,None,False),'C05':(2,'Geometry study not yet done.',False),'C06':(2,None,False),
 'C07':(4,'Ground optical-clock chronometric geodesy demonstrated; ACES (2025) in orbit at microwave-clock stability.',True),
 'C08':(5,'Ground atom gradiometers exist; NASA QGGPf and ESA CARIOQA pathfinders scheduled in this window.',True),
 'C09':(3,'Ground stellar intensity interferometry proves the physics.',False),
 'C10':(2,'Theory and tabletop experiments.',False),
 'C11':(3,'Hunga Tonga (2022) acoustic-gravity waves were seen in airglow from ground and from GOLD/ICON.',False),
 'C12':(3,'Mantle-scale induction from Swarm published.',False),
 'C13':(4,'Lab and shipboard demos.',False),
 'C14':(2,None,False),
 'C15':(3,'ANITA incidental detections serve as proof of concept.',False),
 'C16':(3,'Laboratory demonstrations.',False),
 'C17':(2,'Every component is TRL 9; the application is unvalidated.',False),
 'C18':(4,'Ground Rayleigh–Brillouin lidar demos.',False),
 'C19':(4,'Ground and airborne demos reported; verify.',True),
 'C20':(2,None,False),
 'C21':(3,'Individual-link characterization is the proof of concept.',False),
 'C22':(5,'Airborne LWIR polarimeters.',False),
 'C23':(2,'Lab demos at meter scale. Physics-dubious at orbital ranges.',False),
 'C24':(3,'Airborne SAL (2011) validates optical aperture synthesis; orbital motion compensation unaddressed. Physics-dubious.',False),
}
DUBIOUS = {'C23','C24'}

for t in D['techniques']:
    trl, note, ver = T[t['id']]
    t['trl'] = trl; t['trlNote'] = note or ''; t['verify'] = ver
    t.pop('status', None)
for c in D['concepts']:
    trl, note, ver = C[c['id']]
    c['trl'] = trl; c['trlNote'] = note or ''; c['verify'] = ver
    c['dubious'] = c['id'] in DUBIOUS
    c.pop('status', None)

D.pop('statuses', None)
D['trlBands'] = [
  {'id':'b9','name':'TRL 9, mission-proven','lo':9,'hi':9},
  {'id':'b78','name':'TRL 7–8, flight-demonstrated','lo':7,'hi':8},
  {'id':'b46','name':'TRL 4–6, validated on ground or aircraft','lo':4,'hi':6},
  {'id':'b13','name':'TRL 1–3, concept and proof of principle','lo':1,'hi':3},
]
D['trlScale'] = [
  (1,'Basic principles observed and reported'),
  (2,'Technology concept and application formulated'),
  (3,'Analytical and experimental proof of concept'),
  (4,'Component or breadboard validated in laboratory'),
  (5,'Component or breadboard validated in relevant environment'),
  (6,'System or subsystem model or prototype demonstrated in relevant environment'),
  (7,'System prototype demonstrated in space environment'),
  (8,'Actual system completed and flight-qualified through test and demonstration'),
  (9,'Actual system proven through successful mission operations'),
]
D['meta']['date'] = '2026-09-13'
D['meta']['updated'] = '2026-10-13'
D['meta']['caveat'] = ('TRL values are my estimates against NASA definitions, judged for use from space in the stated application, '
  'from knowledge through mid-2026. They are hypotheses to verify, and the TRL of a technique is not the TRL of any specific instrument. '
  'Items marked verify sit in the 2025–2027 pathfinder window where a single launch moves the number. Novelty claims remain the camouflaged failure mode.')
D['meta']['subtitle'] = 'A source × interaction grid of every established technique, the empty cells, and the concepts that could fill them, with technology readiness per entry'
json.dump(D, open('/home/claude/data.json','w'), ensure_ascii=False, indent=1)
print('ok', sum(1 for t in D['techniques'] if t['verify']), 'techniques and', sum(1 for c in D['concepts'] if c['verify']), 'concepts flagged verify')
