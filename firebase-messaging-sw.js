// خدمة الإشعارات في الخلفية (Firebase Cloud Messaging) — لازم يكون اسم
// الملف ده "firebase-messaging-sw.js" بالظبط وموجود جنب index.html في نفس
// مجلد الاستضافة (نفس المستوى)، عشان المتصفح يقدر يلاقيه.

importScripts('https://www.gstatic.com/firebasejs/10.7.1/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.7.1/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "AIzaSyBelPbcjbCtHTuCDIKYqp_SHPIWIkb9w3M",
  authDomain: "chillout-eeb6e.firebaseapp.com",
  projectId: "chillout-eeb6e",
  storageBucket: "chillout-eeb6e.firebasestorage.app",
  messagingSenderId: "474818688224",
  appId: "1:474818688224:web:967d54f5c0769001690c85"
});

var messaging = firebase.messaging();
/* عشان أي تحديث جديد للملف ده يشتغل فورًا على جهاز المستخدم من غير ما يحتاج يقفل التطبيق تمامًا ويفتحه تاني */
self.addEventListener('install', function (event) { self.skipWaiting(); });
self.addEventListener('activate', function (event) { event.waitUntil(self.clients.claim()); });

// خريطة الأيقونة والـ tag لكل نوع إشعار — علشان كل نوع يظهر بأيقونة مختلفة
// على شاشة الموبايل (مش نفس الأيقونة الافتراضية للكل).
// - icon: الأيقونة الكبيرة اللي بتظهر جنب النص في الإشعار
// - badge: الأيقونة الصغيرة (monochrome) في status bar على أندرويد
// - tag: علشان الـ OS يجمع الإشعارات من نفس النوع تحت مجموعة واحدة
// - vibrate: نمط اهتزاز مختلف لكل نوع (اختياري)
var TOPIC_STYLES = {
  order:    { icon: 'icon-order.png',    badge: 'icon-order.png',    tag: 'hhm-order',    vibrate: [200, 100, 200] },
  overdue:  { icon: 'icon-overdue.png',  badge: 'icon-overdue.png',  tag: 'hhm-overdue',  vibrate: [300, 200, 300, 200, 300] },
  checks:   { icon: 'icon-checks.png',   badge: 'icon-checks.png',   tag: 'hhm-checks',   vibrate: [200, 100, 200] },
  salary:   { icon: 'icon-salary.png',   badge: 'icon-salary.png',   tag: 'hhm-salary',   vibrate: [200] },
  request:  { icon: 'icon-192.png',      badge: 'icon-192.png',      tag: 'hhm-request',  vibrate: [200, 100, 200] }
};

// كاش صغير لمنع تكرار نفس الإشعار خلال 60 ثانية — لو نفس الرسالة وصلت
// مرتين من السيرفر أو من foreground + background.
var __swSeen = {};

// بيستقبل الإشعار وقت ما التطبيق مقفول أو في الخلفية، ويعرضه كإشعار نظام
// عادي على الموبايل — بأيقونة/لون مختلف لكل نوع إشعار (شيكات/متأخرات/أوامر شراء).
messaging.onBackgroundMessage(function (payload) {
  var d = payload.data || {};
  var n = payload.notification || {};
  var title = d.title || n.title || 'إشعار جديد';
  var body  = d.body  || n.body  || '';
  var topic = (d.topic || d.kind || '').toLowerCase();

  // ضم الكمية للنص لو مبعوتة (مثل "20 طن" لأمر شراء)
  if (d.qty && String(body).indexOf(d.qty) === -1) {
    var unit = d.unit ? (' ' + d.unit) : '';
    body = body + (body ? ' — ' : '') + d.qty + unit;
  }

  // دِيدَاب: نفس النص خلال 60 ثانية بيتشال.
  var key = String(title) + '||' + String(body);
  var now = Date.now();
  if (__swSeen[key] && (now - __swSeen[key]) < 60000) return;
  __swSeen[key] = now;
  // نظف الكاش من العناصر القديمة (أكتر من 5 دقايق) عشان مايكبرش على الفاضي
  Object.keys(__swSeen).forEach(function (k) {
    if (now - __swSeen[k] > 300000) delete __swSeen[k];
  });

  var style = TOPIC_STYLES[topic] || TOPIC_STYLES.request;
  var options = {
    body: body,
    icon: style.icon,
    badge: style.badge,
    tag:  style.tag + '-' + (d.dedupId || key),
    renotify: true,
    vibrate: style.vibrate,
    data: Object.assign({}, d, { topic: topic, ts: now })
  };
  self.registration.showNotification(title, options);
});

// لو المستخدم داس على الإشعار: يفتح/يركّز نافذة التطبيق، ولو الإشعار له action معينة
// (زي الشيكات/المتأخرات/أوامر الشراء) يفتح الشاشة المناسبة مباشرة.
self.addEventListener('notificationclick', function (event) {
  event.notification.close();
  var data = event.notification.data || {};
  var targetUrl = self.registration.scope;
  var actionParamMap = {
    checksReport:    'openChecksReport',
    overdueReport:   'openOverdueReport',
    orderSubmit:     'openOrderForm',
    mySalary:        'openMySalary',
    pendingRequests: 'openPendingRequests'
  };
  var actionParam = actionParamMap[data.action];
  if (actionParam) {
    targetUrl += (targetUrl.indexOf('?') === -1 ? '?' : '&') + actionParam + '=1';
  }
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (clientList) {
      for (var i = 0; i < clientList.length; i++) {
        var client = clientList[i];
        if ('focus' in client) {
          client.focus();
          if ('navigate' in client) { client.navigate(targetUrl); }
          return;
        }
      }
      if (clients.openWindow) { return clients.openWindow(targetUrl); }
    })
  );
});
