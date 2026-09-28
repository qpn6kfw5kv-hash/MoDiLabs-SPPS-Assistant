(function() {
  'use strict';

  var INTRO_KEY = 'spps-intro-played-this-session';
  var container = null;
  var canvas = null;
  var renderer = null;
  var scene = null;
  var camera = null;
  var THREE = null;
  var peptideGroups = [];
  var materials = {};
  var atomGeometry = null;
  var bondGeometry = null;
  var fieldTime = 0;
  var mouse = { x: 0, y: 0, targetX: 0, targetY: 0 };
  var lastFrameTime = null;
  var reducedMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var diagnostics = {
    mode: 'pending',
    peptideCount: 0,
    frameCount: 0,
    lastRenderAt: 0,
    pixelCheck: null
  };
  window.__sppsPeptideBackgroundDiagnostics = diagnostics;
  document.documentElement.dataset.peptideBackgroundScript = 'loaded';

  var molecularModels = window.SPPS_PEPTIDE_MODELS || [];

  function randBetween(min, max) {
    return min + Math.random() * (max - min);
  }

  function viewExtent(z) {
    var halfHeight = Math.tan(THREE.MathUtils.degToRad(camera.fov / 2)) * Math.max(1, camera.position.z - z);
    return { x: halfHeight * camera.aspect, y: halfHeight };
  }

  function generatePeptideConfigs() {
    var count = window.innerWidth >= 920 ? 4 : 3;
    var sectors = [[-0.7, 0.48], [0.6, -0.52], [0.45, 0.68], [-0.55, -0.65]];
    var depths = [3.5, -8, -24, -40];
    var headings = [2.6, -0.55, -2.2, 0.85];
    return sectors.slice(0, count).map(function(sector, index) {
      var z = depths[index] + randBetween(-1, 1);
      var extent = viewExtent(z);
      var angle = headings[index] + randBetween(-0.2, 0.2);
      var speed = randBetween(0.055, 0.085) * (camera.position.z - z) / 18;
      return {
        model: molecularModels[index % molecularModels.length],
        position: [sector[0] * extent.x, sector[1] * extent.y, z],
        velocity: [Math.cos(angle) * speed, Math.sin(angle) * speed, (index % 2 ? 1 : -1) * randBetween(0.045, 0.085)],
        rotation: [randBetween(-0.7, 0.7), randBetween(-0.7, 0.7), randBetween(-Math.PI, Math.PI)],
        spin: [randBetween(-0.009, 0.009), randBetween(-0.012, 0.012), randBetween(-0.006, 0.006)],
        scale: window.innerWidth < 760 ? 0.52 : 0.76,
        atomScale: 1,
        phase: randBetween(0, Math.PI * 2)
      };
    });
  }

  function publishDiagnostics() {
    var pixelCheck = diagnostics.pixelCheck || {};
    document.documentElement.dataset.peptideBackgroundMode = diagnostics.mode;
    document.documentElement.dataset.peptideBackgroundFrames = String(diagnostics.frameCount);
    if (!container) return;
    container.dataset.script = 'loaded';
    container.dataset.mode = diagnostics.mode;
    container.dataset.peptideCount = String(diagnostics.peptideCount);
    container.dataset.frameCount = String(diagnostics.frameCount);
    container.dataset.pixelColored = pixelCheck.coloredPixels == null ? '' : String(pixelCheck.coloredPixels);
    container.dataset.pixelRatio = pixelCheck.nonTransparentRatio == null ? '' : String(pixelCheck.nonTransparentRatio);
    diagnostics.motion = peptideGroups.map(function(group) { return { position: group.position.toArray(), velocity: group.userData.velocity.toArray() }; });
    diagnostics.drawCalls = renderer ? renderer.info.render.calls : 0;
    if (peptideGroups.length) {
      container.dataset.sequenceLengths = peptideGroups.map(function(group) {
        return String(group.userData.sequenceLength || 0);
      }).join(',');
      container.dataset.flowSample = peptideGroups.slice(0, 3).map(function(group) {
        return [
          group.position.x.toFixed(2),
          group.position.y.toFixed(2),
          group.position.z.toFixed(2)
        ].join(',');
      }).join('|');
    }
  }

  function runIntro() {
    var intro = document.getElementById('site-intro');
    if (!intro) return;

    var shouldPlay = false;
    try {
      shouldPlay = !sessionStorage.getItem(INTRO_KEY);
      if (shouldPlay) sessionStorage.setItem(INTRO_KEY, '1');
    } catch (err) {
      shouldPlay = !window.__sppsIntroPlayed;
      window.__sppsIntroPlayed = true;
    }

    if (!shouldPlay || reducedMotion) {
      diagnostics.intro = 'skipped';
      publishDiagnostics();
      intro.classList.add('is-skipped');
      return;
    }

    diagnostics.intro = 'played';
    publishDiagnostics();
    document.body.classList.add('intro-running');
    window.setTimeout(function() {
      document.body.classList.remove('intro-running');
      intro.classList.add('is-complete');
      diagnostics.intro = 'complete';
      publishDiagnostics();
    }, 1500);
  }

  function getTheme() {
    return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
  }

  function themePalette() {
    var dark = getTheme() === 'dark';
    return {
      carbon: dark ? 0x737d88 : 0x303640,
      nitrogen: dark ? 0x3974ff : 0x1d4ed8,
      oxygen: dark ? 0xef4444 : 0xdc2626,
      sulfur: dark ? 0xf4c430 : 0xd99a12,
      hydrogen: dark ? 0xffffff : 0xf7f7f2,
      bond: dark ? 0xe6edf5 : 0xd8dee6,
      emissive: dark ? 0x0b1722 : 0xf2f8ff,
      atomOpacity: 1,
      hydrogenOpacity: 1,
      bondOpacity: 1
    };
  }

  async function loadThree() {
    if (window.THREE) return window.THREE;
    try {
      return await import('https://cdn.jsdelivr.net/npm/three@0.160.0/build/three.module.js');
    } catch (err) {
      console.warn('Three.js non disponibile per lo sfondo peptidico', err);
      return null;
    }
  }

  function createMaterials() {
    var palette = themePalette();
    var Material = THREE.MeshPhysicalMaterial || THREE.MeshStandardMaterial;
    materials = {
      C: new Material({ color: palette.carbon, roughness: 0.3, metalness: 0.04, clearcoat: 0.55, clearcoatRoughness: 0.22, transparent: false, opacity: palette.atomOpacity }),
      N: new Material({ color: palette.nitrogen, roughness: 0.28, metalness: 0.04, clearcoat: 0.52, clearcoatRoughness: 0.2, transparent: false, opacity: palette.atomOpacity }),
      O: new Material({ color: palette.oxygen, roughness: 0.28, metalness: 0.04, clearcoat: 0.52, clearcoatRoughness: 0.2, transparent: false, opacity: palette.atomOpacity }),
      S: new Material({ color: palette.sulfur, roughness: 0.25, metalness: 0.04, clearcoat: 0.58, clearcoatRoughness: 0.18, transparent: false, opacity: palette.atomOpacity }),
      H: new Material({ color: palette.hydrogen, roughness: 0.22, metalness: 0.02, clearcoat: 0.45, clearcoatRoughness: 0.18, transparent: false, opacity: palette.hydrogenOpacity }),
      bond: new Material({ color: palette.bond, roughness: 0.34, metalness: 0.02, clearcoat: 0.34, clearcoatRoughness: 0.25, transparent: false, opacity: palette.bondOpacity })
    };
  }

  function updateMaterials() {
    var palette = themePalette();
    if (!materials.C) return;
    if (scene && scene.fog) scene.fog.color.setHex(getTheme() === 'dark' ? 0x071018 : 0xf8fafc);
    materials.C.color.setHex(palette.carbon);
    materials.N.color.setHex(palette.nitrogen);
    materials.O.color.setHex(palette.oxygen);
    materials.S.color.setHex(palette.sulfur);
    materials.H.color.setHex(palette.hydrogen);
    materials.bond.color.setHex(palette.bond);
    materials.C.opacity = palette.atomOpacity;
    materials.N.opacity = palette.atomOpacity;
    materials.O.opacity = palette.atomOpacity;
    materials.S.opacity = palette.atomOpacity;
    materials.H.opacity = palette.hydrogenOpacity;
    materials.bond.opacity = palette.bondOpacity;
  }

  function vec(arr) {
    return new THREE.Vector3(arr[0], arr[1], arr[2]);
  }

  function addAtom(group, name, element, position, scale) {
    var radius = { C: 0.14, N: 0.14, O: 0.135, S: 0.17, H: 0.085 }[element] || 0.14;
    var geometry = atomGeometry;

    var mesh = new THREE.Mesh(geometry, materials[element] || materials.C);
    mesh.position.copy(position);
    mesh.scale.setScalar(radius * scale);
    mesh.userData.element = element;
    group.add(mesh);
    group.userData.atoms[name] = mesh;
    return mesh;
  }

  function addBond(group, atomA, atomB, scale, order) {
    var a = group.userData.atoms[atomA];
    var b = group.userData.atoms[atomB];
    if (!a || !b) return;
    var direction = b.position.clone().sub(a.position);
    var length = direction.length();
    var unit = direction.clone().normalize();
    var offset = new THREE.Vector3().crossVectors(unit, Math.abs(unit.y) < 0.8
      ? new THREE.Vector3(0, 1, 0) : new THREE.Vector3(1, 0, 0)).normalize().multiplyScalar(0.043);
    var hydrogen = a.userData.element === 'H' || b.userData.element === 'H';
    for (var line = 0; line < (order || 1); line++) {
      var mesh = new THREE.Mesh(bondGeometry, materials.bond);
      mesh.position.copy(a.position).add(b.position).multiplyScalar(0.5);
      if (order === 2) mesh.position.addScaledVector(offset, line === 0 ? -1 : 1);
      mesh.quaternion.setFromUnitVectors(new THREE.Vector3(0, 1, 0), unit);
      var radius = (hydrogen ? 0.022 : order === 2 ? 0.025 : 0.036) * scale;
      mesh.scale.set(radius, length, radius);
      group.add(mesh);
    }
  }

  function batchMolecule(group) {
    // Shared geometry and instancing keep complete all-atom models inexpensive.
    var batches = {};
    group.children.forEach(function(mesh) {
      var key = mesh.geometry.id + ':' + mesh.material.id;
      if (!batches[key]) batches[key] = [];
      batches[key].push(mesh);
    });
    group.clear();
    Object.keys(batches).forEach(function(key) {
      var meshes = batches[key];
      var batch = new THREE.InstancedMesh(meshes[0].geometry, meshes[0].material, meshes.length);
      meshes.forEach(function(mesh, index) {
        mesh.updateMatrix();
        batch.setMatrixAt(index, mesh.matrix);
      });
      batch.computeBoundingSphere();
      group.add(batch);
    });
  }

  function buildPeptide(config, index) {
    var group = new THREE.Group();
    group.userData.atoms = {};
    var model = config.model;
    var atomScale = config.atomScale || 1;
    var points = model.atoms.map(function(atom) { return vec(atom.slice(2)); });
    var center = new THREE.Box3().setFromPoints(points).getCenter(new THREE.Vector3());
    var alphaCarbons = model.atoms.filter(function(atom) { return /_CA$/.test(atom[0]); });
    var axis = vec(alphaCarbons[alphaCarbons.length - 1].slice(2))
      .sub(vec(alphaCarbons[0].slice(2))).normalize();
    var orientation = new THREE.Quaternion().setFromUnitVectors(axis, new THREE.Vector3(1, 0, 0));

    // Rigid rotation and uniform scale preserve every deposited bond angle.
    model.atoms.forEach(function(atom, i) {
      var position = points[i].clone().sub(center).applyQuaternion(orientation).multiplyScalar(0.42);
      addAtom(group, atom[0], atom[1], position, atomScale);
    });
    model.bonds.forEach(function(bond) {
      addBond(group, bond[0], bond[1], atomScale, bond[2] || 1);
    });
    batchMolecule(group);
    group.userData.radius = Math.max.apply(null, Object.values(group.userData.atoms).map(function(atom) { return atom.position.length(); })) * config.scale + 0.2;
    group.userData.source = model.source + ':' + model.chain + ':' + model.residues.join('-');

    group.position.set(config.position[0], config.position[1], config.position[2]);
    group.rotation.set(config.rotation[0], config.rotation[1], config.rotation[2]);
    group.scale.setScalar(config.scale || 1);
    group.userData.sequenceLength = model.sequence.length;
    group.userData.baseScale = config.scale;
    group.userData.flowPosition = vec(config.position);
    group.userData.velocity = vec(config.velocity);
    group.userData.targetVelocity = vec(config.velocity);
    group.userData.spin = vec(config.spin);
    group.userData.nextTurn = randBetween(24, 48);
    group.userData.phase = config.phase;
    return group;
  }

  function continueBeyondFrame(group) {
    var data = group.userData;
    var p = data.flowPosition;
    var extent = viewExtent(p.z);
    var radius = data.radius;
    var behind = p.z - radius > camera.position.z;
    var offscreen = Math.abs(p.x) > extent.x + radius * 1.5 || Math.abs(p.y) > extent.y + radius * 1.5;
    if (!behind && !offscreen) return;
    // Recycle only outside the entire camera frustum; no visible bounce or reset.
    var z = randBetween(-44, 3);
    extent = viewExtent(z);
    var side = Math.floor(randBetween(0, 4));
    var horizontal = side < 2;
    var sign = side % 2 ? -1 : 1;
    p.set(horizontal ? sign * (extent.x + radius * 1.1) : randBetween(-extent.x, extent.x),
      horizontal ? randBetween(-extent.y, extent.y) : sign * (extent.y + radius * 1.1), z);
    var speed = randBetween(0.06, 0.1) * (camera.position.z - z) / 18;
    data.velocity.set(horizontal ? -sign * speed : randBetween(-speed, speed),
      horizontal ? randBetween(-speed, speed) : -sign * speed, randBetween(-0.08, 0.08));
    data.targetVelocity.copy(data.velocity);
    data.nextTurn = fieldTime + randBetween(30, 60);
  }

  function advancePeptide(group, delta) {
    var data = group.userData;
    if (fieldTime > data.nextTurn) {
      var angle = Math.atan2(data.targetVelocity.y, data.targetVelocity.x) + randBetween(-0.65, 0.65);
      var speed = Math.hypot(data.targetVelocity.x, data.targetVelocity.y);
      data.targetVelocity.set(Math.cos(angle) * speed, Math.sin(angle) * speed, randBetween(-0.085, 0.085));
      data.nextTurn = fieldTime + randBetween(24, 55);
    }
    data.velocity.lerp(data.targetVelocity, 1 - Math.exp(-delta * 0.12));
    data.flowPosition.addScaledVector(data.velocity, delta);
    continueBeyondFrame(group);
    group.position.copy(data.flowPosition);
    var parallax = 0.12 * 18 / Math.max(6, camera.position.z - group.position.z);
    group.position.x += mouse.x * parallax;
    group.position.y -= mouse.y * parallax;
    group.rotation.x += data.spin.x * delta;
    group.rotation.y += data.spin.y * delta;
    group.rotation.z += data.spin.z * delta;
  }

  function resize() {
    if (!renderer || !camera) return;
    var width = window.innerWidth || 1;
    var height = window.innerHeight || 1;
    renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 1.7));
    renderer.setSize(width, height, false);
    camera.aspect = width / height;
    camera.position.z = width < 760 ? 13.5 : 11.5;
    camera.updateProjectionMatrix();

  }

  function animate(time) {
    if (!renderer || !scene || !camera) return;
    var seconds = (time || 0) * 0.001;
    var delta = lastFrameTime == null ? 0.016 : Math.min(0.06, Math.max(0.001, seconds - lastFrameTime));
    lastFrameTime = seconds;
    var pointerBlend = 1 - Math.exp(-delta * 2.8);
    mouse.x += (mouse.targetX - mouse.x) * pointerBlend;
    mouse.y += (mouse.targetY - mouse.y) * pointerBlend;

    fieldTime += delta;
    peptideGroups.forEach(function(group) { advancePeptide(group, delta); });

    renderer.render(scene, camera);
    diagnostics.frameCount += 1;
    diagnostics.lastRenderAt = Date.now();
    if (diagnostics.frameCount <= 3) updatePixelDiagnostics();
    if (diagnostics.frameCount % 60 === 0) publishDiagnostics();
    window.requestAnimationFrame(animate);
  }

  function updatePixelDiagnostics() {
    if (!renderer) return;
    try {
      var gl = renderer.getContext();
      var width = gl.drawingBufferWidth;
      var height = gl.drawingBufferHeight;
      if (!width || !height || width * height > 3000000) {
        diagnostics.pixelCheck = { checked: false, reason: 'buffer too large', width: width, height: height };
        publishDiagnostics();
        return;
      }
      var pixels = new Uint8Array(width * height * 4);
      gl.readPixels(0, 0, width, height, gl.RGBA, gl.UNSIGNED_BYTE, pixels);
      var nonTransparent = 0;
      var colored = 0;
      for (var i = 0; i < pixels.length; i += 4) {
        if (pixels[i + 3] > 0) nonTransparent += 1;
        if (pixels[i] + pixels[i + 1] + pixels[i + 2] > 12) colored += 1;
      }
      diagnostics.pixelCheck = {
        checked: true,
        width: width,
        height: height,
        nonTransparentPixels: nonTransparent,
        coloredPixels: colored,
        nonTransparentRatio: Number((nonTransparent / (width * height)).toFixed(5))
      };
      publishDiagnostics();
    } catch (err) {
      diagnostics.pixelCheck = { checked: false, reason: err && err.message ? err.message : 'readPixels failed' };
      publishDiagnostics();
    }
  }

  async function initThreeBackground() {
    THREE = await loadThree();
    if (!THREE) {
      if (container) container.classList.add('fallback');
      diagnostics.mode = 'fallback';
      publishDiagnostics();
      return false;
    }

    if (!molecularModels.length) throw new Error('Missing experimental peptide models');
    createMaterials();
    atomGeometry = new THREE.SphereGeometry(1, 20, 14);
    bondGeometry = new THREE.CylinderGeometry(1, 1, 1, 10);
    scene = new THREE.Scene();
    camera = new THREE.PerspectiveCamera(36, 1, 0.1, 180);
    camera.position.set(0, 0, 11.5);

    renderer = new THREE.WebGLRenderer({
      canvas: canvas,
      alpha: true,
      antialias: true,
      preserveDrawingBuffer: true,
      powerPreference: 'high-performance'
    });
    renderer.setClearColor(0x000000, 0);
    scene.fog = new THREE.Fog(getTheme() === 'dark' ? 0x071018 : 0xf8fafc, 24, 100);

    scene.add(new THREE.AmbientLight(0xffffff, 0.58));
    var keyLight = new THREE.DirectionalLight(0xffffff, 0.74);
    keyLight.position.set(3, 4, 7);
    scene.add(keyLight);
    var rimLight = new THREE.PointLight(0x5eead4, 1.15, 26);
    rimLight.position.set(-4, 2, 5);
    scene.add(rimLight);

    resize();
    peptideGroups = generatePeptideConfigs().map(function(config, index) {
      var peptide = buildPeptide(config, index);
      scene.add(peptide);
      return peptide;
    });
    diagnostics.mode = 'three';
    diagnostics.peptideCount = peptideGroups.length;
    diagnostics.sources = peptideGroups.map(function(group) { return group.userData.source; });
    publishDiagnostics();

    window.addEventListener('resize', resize);
    window.addEventListener('pointermove', function(event) {
      mouse.targetX = (event.clientX / Math.max(window.innerWidth, 1)) * 2 - 1;
      mouse.targetY = (event.clientY / Math.max(window.innerHeight, 1)) * 2 - 1;
    }, { passive: true });
    window.addEventListener('spps-theme-change', updateMaterials);

    var observer = new MutationObserver(updateMaterials);
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });

    resize();
    if (reducedMotion) {
      renderer.render(scene, camera);
      diagnostics.frameCount = 1;
      diagnostics.lastRenderAt = Date.now();
      updatePixelDiagnostics();
    } else {
      window.requestAnimationFrame(animate);
    }
    return true;
  }

  function boot() {
    container = document.getElementById('peptide-background');
    canvas = document.getElementById('peptide-background-canvas');
    if (!container || !canvas) {
      diagnostics.mode = 'missing-elements';
      publishDiagnostics();
      return;
    }
    publishDiagnostics();
    initThreeBackground().finally(runIntro);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot, { once: true });
  } else {
    boot();
  }
})();
