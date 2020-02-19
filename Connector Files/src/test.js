//import {ondescribe, onexecute} from './index.js';
import test from 'ava';
//import '@k2oss/k2-broker-core/test-framework';

test('foo', t => {
    t.pass();
});

// function mock(name, value) {
//     global[name] = value;
// }
//
// test('describe returns the hardcoded instance', async (t) => {
//     let schema = null;
//     mock('postSchema', function (result) {
//         schema = result;
//     });
//     await Promise.resolve(ondescribe());
//     t.deepEqual(schema, {
//         teams: {
//             "com.k2.todo": {
//                 displayName: "TODO",
//                 description: "Manages a TODO list",
//                 properties: {
//                     "com.k2.todo.id": {
//                         displayName: "ID",
//                         type: "number"
//                     },
//                     "com.k2.todo.userId": {
//                         displayName: "User ID",
//                         type: "number"
//                     },
//                     "com.k2.todo.title": {
//                         displayName: "Title",
//                         type: "string"
//                     },
//                     "com.k2.todo.completed": {
//                         displayName: "Completed",
//                         type: "boolean"
//                     }
//                 },
//                 methods: {
//                     "com.k2.todo.get": {
//                         displayName: "Get TODO",
//                         type: "read",
//                         inputs: ["com.k2.todo.id"],
//                         outputs: ["com.k2.todo.id", "com.k2.todo.userId", "com.k2.todo.title", "com.k2.todo.completed"]
//                     }
//                 }
//             }
//         }
//     });
//     t.pass();
// });

// test('onexecute fails with the invalid object', async (t) => {
//     let obj = 'invalidObject';
//     let error = await t.throwsAsync(Promise.resolve(onexecute(obj, '', {}, {})));
//     t.deepEqual(error.message, `The object ${obj} is not supported.`);
//     t.pass();
// });

// test('team object execute fails with the wrong method', async (t) => {
//     let obj = 'team';
//     let method = 'invalidMethod';
//     let error = await t.throwsAsync(Promise.resolve(onexecute(obj, method, {}, {})));
//     t.deepEqual(error.message, `The method ${method} is not supported.`);
//     t.pass();
// });


// test('execute passes', async (t) => {
//     let xhr = null;
//     class XHR {
//         constructor() {
//             xhr = this.recorder = {};
//             this.recorder.headers = {};
//         }
//         open(method, url) {
//             this.recorder.opened = { method, url };
//         }
//         setRequestHeader(key, value) {
//             this.recorder.headers[key] = value;
//         }
//         send() {
//             queueMicrotask(() => {
//                 this.readyState = 4;
//                 this.status = 200;
//                 this.responseText = JSON.stringify({
//                     "id": 123,
//                     "userId": 51,
//                     "title": "Groceries",
//                     "completed": false
//                 });
//                 this.onreadystatechange();
//                 delete this.responseText;
//             });
//         }
//     }
//     mock('XMLHttpRequest', XHR);
//     let result = null;
//     function pr(r) {
//         result = r;
//     }
//     mock('postResult', pr);
//     await Promise.resolve(onexecute('com.k2.todo', 'com.k2.todo.get', {}, {
//         "com.k2.todo.id": 123
//     }));
//     t.deepEqual(xhr, {
//         opened: {
//             method: 'GET',
//             url: 'https://jsonplaceholder.typicode.com/todos/123'
//         },
//         headers: {
//             'test': 'test value'
//         }
//     });
//     t.deepEqual(result, {
//         "com.k2.todo.id": 123,
//         "com.k2.todo.userId": 51,
//         "com.k2.todo.title": "Groceries",
//         "com.k2.todo.completed": false
//     });
//     t.pass();
// });
