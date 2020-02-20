import test from 'ava';
import '@k2oss/k2-broker-core/test-framework';
import './index';

function mock(name: string, value: any) 
{
    global[name] = value;
}

test('pass', t => {
    t.pass();
});

test('onexecute fails for invalid object', t => {
    t.throws(function () {
        let obj = 'invalidObject';
        onexecute(obj, '', {}, {});
    });
});

const promise = () => Promise.reject(new Error('🦄'));
test('rejects', async t => {
    const error = await t.throwsAsync(promise);
    t.is(error.message, '🦄');
});

test('executeTeamArchive succeeds', async t => {

    let xhr: {[key:string]: any} = null;
    class XHR {
        public onreadystatechange: () => void;
        public readyState: number;
        public status: number;
        public responseText: string;
        private recorder: {[key:string]: any};

        constructor() {
            xhr = this.recorder = {};
            this.recorder.headers = {};
        }

        open(method: string, url: string) {
            this.recorder.opened = {method, url};
        }

        setRequestHeader(key: string, value: string) {
            this.recorder.headers[key] = value;
        }

        send() {
            queueMicrotask(() =>
            {
                this.readyState = 4;
                this.status = 200;
                this.responseText = JSON.stringify({
                    "id": 1234,
                    "requestStatusUrl": undefined,
                    "isSuccessful": true
                });
                this.onreadystatechange();
                delete this.responseText;
            });
        }
    }

    mock('XMLHttpRequest', XHR);

    let result: any = null;
    function pr(r: any) {
        result = r;
    }

    mock('postResult', pr);

    await onexecute('team', 'archive', {}, {"id":1234});

    // t.deepEqual(xhr, {
    //     opened: {
    //         method: 'GET',
    //         url: 'https://jsonplaceholder.typicode.com/todos/123'
    //     },
    //     headers: {
    //         'test': 'test value'
    //     }
    // });

    t.deepEqual(result, {
        "id": 1234,
        "requestStatusUrl": undefined,
        "isSuccessful": true
    });

    t.pass();
});