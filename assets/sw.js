if (!self.define) {
  let c,
    i = {};
  const n = (n, e) => (
    (n = new URL(n + '.js', e).href),
    i[n] ||
      new Promise(i => {
        if ('document' in self) {
          const c = document.createElement('script');
          (c.src = n), (c.onload = i), document.head.appendChild(c);
        } else (c = n), importScripts(n), i();
      }).then(() => {
        let c = i[n];
        if (!c) throw new Error(`Module ${n} didn’t register its module`);
        return c;
      })
  );
  self.define = (e, o) => {
    const f = c || ('document' in self ? document.currentScript.src : '') || location.href;
    if (i[f]) return;
    let r = {};
    const s = c => n(c, f),
      a = { module: { uri: f }, exports: r, require: s };
    i[f] = Promise.all(e.map(c => a[c] || s(c))).then(c => (o(...c), r));
  };
}
define(['./workbox-dae083bf'], function (c) {
  'use strict';
  self.addEventListener('message', c => {
    c.data && 'SKIP_WAITING' === c.data.type && self.skipWaiting();
  }),
    c.precacheAndRoute(
      [
        { url: 'favicon.png', revision: '1bc73a93f5aa3c02fe25899cf0a2e10d' },
        { url: 'graff.png', revision: '374e94d8e3aed9fa88a9923e5eaacea2' },
        { url: 'icons/100x100-icon.png', revision: 'fba112b56be61bcb1ade15c2f96fe129' },
        { url: 'icons/1024x1024-icon.png', revision: 'b1cfd1dd029d2fbde88ff06c573dc290' },
        { url: 'icons/107x107-icon.png', revision: '453bc0204d75d8cb40cb0a76fc7f1ebb' },
        { url: 'icons/114x114-icon.png', revision: '3b5bb013a37d95a821b5ff6d75ce8c4d' },
        { url: 'icons/120x120-icon.png', revision: '4c02e76958a1b90df8d1cdfad975947c' },
        { url: 'icons/1240x1240-icon.png', revision: 'e92613d2884f1400385511323c6345c6' },
        { url: 'icons/1240x600-icon.png', revision: '51e1d4db91af3161abc10e12cf4bb2b2' },
        { url: 'icons/128x128-icon.png', revision: 'f61c9e5c119ca1c392eaffb31508e27e' },
        { url: 'icons/142x142-icon.png', revision: 'b16b5ceeb64d483772d26b26114008d4' },
        { url: 'icons/144x144-icon.png', revision: 'c3b524ba22a3d93a6beddcaec829a29d' },
        { url: 'icons/150x150-icon.png', revision: 'a44d2c170d7433bd5c5f703ec2022026' },
        { url: 'icons/152x152-icon.png', revision: '3f0560c65d1821477a275f43b1c34257' },
        { url: 'icons/167x167-icon.png', revision: '3111b534fbb2ab41cc34756f26fc24c2' },
        { url: 'icons/16x16-icon.png', revision: 'c2120634e047ec9df15774ab199cc39f' },
        { url: 'icons/176x176-icon.png', revision: '92a22188b7c07bca96da0bb958bf545c' },
        { url: 'icons/180x180-icon.png', revision: '9a837d241cf7255ae161eadb106e197a' },
        { url: 'icons/188x188-icon.png', revision: '79be323154a83f42321bbab31899a968' },
        { url: 'icons/192x192-icon.png', revision: 'ee95393c980f31f9c75c9309f47afbf9' },
        { url: 'icons/200x200-icon.png', revision: 'ba5e3dc7d2c68f722f5655bf81142267' },
        { url: 'icons/20x20-icon.png', revision: 'afa27e4ba7e2853586fab0386faaf8e1' },
        { url: 'icons/225x225-icon.png', revision: '61df6cc438afd2c2e71294af23da3df2' },
        { url: 'icons/2480x1200-icon.png', revision: '07a59eda0d953ee2f5ea811bf7743e05' },
        { url: 'icons/24x24-icon.png', revision: 'a45a12126bf32c7c816f83377b5d8415' },
        { url: 'icons/256x256-icon.png', revision: '0fd3e9e1243cf36701a503d1e595185e' },
        { url: 'icons/284x284-icon.png', revision: '5a51a5f8a4006a2f88cce7ef5deb13d0' },
        { url: 'icons/29x29-icon.png', revision: '8d7727318e934ee05139c3e6b3be0e43' },
        { url: 'icons/300x300-icon.png', revision: '9f19330173963838f0f7e7eaf8846782' },
        { url: 'icons/30x30-icon.png', revision: '77434c90c0a6cb5404827a40496e273e' },
        { url: 'icons/310x150-icon.png', revision: 'ee45699da5bd51eb76732ba668efcdc7' },
        { url: 'icons/310x310-icon.png', revision: 'ef1ca989c6a4c08720493d7cee8eccf9' },
        { url: 'icons/32x32-icon.png', revision: '47c6ddde2ec12df316cc6d5e791586c7' },
        { url: 'icons/36x36-icon.png', revision: '06f7cb658ba2edc48ec2f9d7d50e11d7' },
        { url: 'icons/388x188-icon.png', revision: '98530c4c607935752882999ef02d01df' },
        { url: 'icons/388x388-icon.png', revision: 'ffbc180819af25ef08b7a9c3d6758a79' },
        { url: 'icons/40x40-icon.png', revision: 'fbcdc47922ab798c71c24ecb6a8b3dfb' },
        { url: 'icons/44x44-icon.png', revision: 'c39c3e945a8b3b34b3b837b19bd07c12' },
        { url: 'icons/465x225-icon.png', revision: '24b57a4818fd62340872c9df2fec82ed' },
        { url: 'icons/465x465-icon.png', revision: '1b129f103a310f1458bc8b5e5fdf4e00' },
        { url: 'icons/48x48-icon.png', revision: '3d90315b767301054a7a3e228efe9105' },
        { url: 'icons/50x50-icon.png', revision: '8c5eddd92388b2f280816c1595a65950' },
        { url: 'icons/512x512-icon.png', revision: 'c242712c53c95f38836a1b40812af3aa' },
        { url: 'icons/55x55-icon.png', revision: '5ed43d66499cb07ae94160aecd9bf1f5' },
        { url: 'icons/57x57-icon.png', revision: 'b31beb0cc36e820116b420f7ef84f795' },
        { url: 'icons/58x58-icon.png', revision: '99f10b080b3b5ea158616d7531fae553' },
        { url: 'icons/600x600-icon.png', revision: 'cfeb0f07ecf2ee78977f599524a4515d' },
        { url: 'icons/60x60-icon.png', revision: '1f185c39ac45ccb9ba6d66e9875b00d7' },
        { url: 'icons/620x300-icon.png', revision: '44e57502ced6f98887d59ec9e5672f67' },
        { url: 'icons/620x620-icon.png', revision: 'd58422e5c2f20eda7e12f3134a7941fa' },
        { url: 'icons/63x63-icon.png', revision: '5e380bb4eaa12732b9f929f5f69e57f8' },
        { url: 'icons/64x64-icon.png', revision: '74953c656daf397090c44d07f8f15886' },
        { url: 'icons/66x66-icon.png', revision: '0b457a847d797d66e365366fa6a3301a' },
        { url: 'icons/71x71-icon.png', revision: 'cedc51d2407bb4aa8cfc45f58b9ad094' },
        { url: 'icons/72x72-icon.png', revision: '4de0237f1f1595953c0c4ae5ae9f8d6b' },
        { url: 'icons/75x75-icon.png', revision: 'bbfa26dad3a24c005a5ab660829c6acf' },
        { url: 'icons/76x76-icon.png', revision: '3ef6ce628a1702f7c8678c8e92d7bba3' },
        { url: 'icons/775x375-icon.png', revision: 'd4fa10700a75e76fbae0888e9486ea34' },
        { url: 'icons/80x80-icon.png', revision: '8ca6e67f0ec1269b1d8eb2febfee8654' },
        { url: 'icons/87x87-icon.png', revision: '8c91cd248829c97ba5e5e93f017a4768' },
        { url: 'icons/88x88-icon.png', revision: '572b1d173ad797e211695488f86d79a9' },
        { url: 'icons/89x89-icon.png', revision: 'f73fef6f9c9955233c6a1f9eb213ddef' },
        { url: 'icons/930x450-icon.png', revision: '2030d1bb6d881c8ede242c471758bff3' },
        { url: 'icons/96x96-icon.png', revision: '521fb6b0d2e5e477b71c5dac1b022b9c' },
        { url: 'manifest.json', revision: 'ee016e1fe5c51f6cca2a570d121c33b8' },
        { url: 'mgt.png', revision: '0a1f53f06c9711cf7d83128cd76e7f8e' },
        { url: 'mgt.storybook.js', revision: 'bebf2959d77ab8a92a7afa4f4780472a' }
      ],
      { ignoreURLParametersMatching: [/^utm_/, /^fbclid$/] }
    );
});
