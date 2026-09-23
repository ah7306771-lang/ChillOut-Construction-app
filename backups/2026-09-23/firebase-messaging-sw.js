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

var messaging = firebase.messaging(); /* عشان أي تحديث جديد للملف ده (زي إضافة خاصية فتح تقرير الشيكات) يشتغل فورًا على جهاز المستخدم من غير ما يحتاج يقفل التطبيق تمامًا ويفتحه تاني */ self.addEventListener('install', function (event) { self.skipWaiting(); }); self.addEventListener('activate', function (event) { event.waitUntil(self.clients.claim()); });

// بيستقبل الإشعار وقت ما التطبيق مقفول أو في الخلفية، ويعرضه كإشعار نظام
// عادي على الموبايل (زي أي إشعار تطبيق تاني).
// ملحوظة: بنقرا العنوان والنص من payload.data (مش payload.notification)
// عشان ده اللي شكل الباك إند بيبعته دلوقتي — ده اللي بيمنع ظهور نفس
// الإشعار مرتين على بعض الأجهزة (أندرويد بالذات).
messaging.onBackgroundMessage(function (payload) {
  var title = (payload.data && payload.data.title) || (payload.notification && payload.notification.title) || 'إشعار جديد';
  var options = {
    body: (payload.data && payload.data.body) || (payload.notification && payload.notification.body) || '',
    icon: 'icon-192.png',
    data: payload.data || {}
  };
  self.registration.showNotification(title, options);
});

// لو المستخدم داس على الإشعار: يفتح/يركّز نافذة التطبيق، ولو الإشعار له action معينة (زي الشيكات) يفتح شاشة تقرير الشيكات مباشرة
self.addEventListener('notificationclick', function (event) {
  event.notification.close();
  var data = event.notification.data || {};
  var targetUrl = self.registration.scope;
  var actionParamMap = {
    checksReport: 'openChecksReport',
    mySalary: 'openMySalary',
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
