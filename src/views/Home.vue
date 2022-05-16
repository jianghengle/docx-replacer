<template>
  <div class="container">
    <section class="section">
      <div v-if="error" class="notification is-danger is-light">
        <button class="delete" @click="error=''"></button>
        {{error}}
      </div>

      <div class="field">
        <label class="label">Replacements</label>

        <div class="field has-addons" v-for="(r, i) in replacements" :key="'r'+i">
          <p class="control">
            <input class="input" type="text" placeholder="name" v-model="r.name">
          </p>
          <p class="control is-expanded">
           <input class="input" type="text" placeholder="value" v-model="r.value">
          </p>
          <p class="control">
            <a class="button" @click="removeReplacements(i)">
              <span class="icon is-small">
                <i class="fas fa-times"></i>
              </span>
            </a>
          </p>
        </div>
        

      </div>

      <div>
        <div class="field is-grouped is-pulled-right">
          <div class="control">
            <button class="button" @click="clearValues">Clear Values</button>
          </div>
          <div class="control">
            <button class="button" @click="removeReplacements(null)">Remove All Replacements</button>
          </div>
        </div>
        <div class="field is-grouped">
          <div class="control">
            <button class="button" @click="addReplacement">
              <span class="icon is-small">
                <i class="fas fa-plus"></i>
              </span>
              <span>Add Replacement</span>
            </button>
          </div>
        </div>

        <div class="field is-grouped">
          <div class="control">
            <button class="button is-link" @click="saveReplacements">Save Replacements</button>
          </div>
        </div>
      </div>

      <hr />

      <div class="field">
        <label class="label">Docx Files</label>

        <div class="field has-addons">
          <div class="control">
            <div class="file has-name">
              <label class="file-label">
                <input class="file-input" type="file" accept=".docx" multiple @change="onFileChange">
                <span class="file-cta">
                  <span class="file-icon">
                    <i class="fas fa-upload"></i>
                  </span>
                  <span class="file-label">
                    Files
                  </span>
                </span>
                <span class="file-name">
                  <span v-if="files && files.length">
                    Selected <strong>{{files.length}}</strong> files
                  </span>
                  <span v-else>Please select files...</span>
                </span>
              </label>
            </div>
          </div>
        </div>

        <div class="control pl-5">
          <ol>
            <li v-for="(f, i) in files" :key="'f-'+i">
              {{f.name}}
            </li>
          </ol>
        </div>

      </div>

      <hr />

      <div class="field is-grouped">
        <div class="control">
          <button class="button is-link" :class="{'is-loading': generating}" @click="generate" :disabled="!canGenerate">Generate Files</button>
        </div>
      </div>

      <div class="modal" :class="{'is-active': saveModal.opened}">
        <div class="modal-background"></div>
        <div class="modal-content">
          <article class="message is-success">
            <div class="message-header">
              <p>Success</p>
              <button class="delete" aria-label="delete" @click="saveModal.opened = false"></button>
            </div>
            <div class="message-body">
              <p>Your replacements have been saved in your current browser's storage.</p> <br/>
              <p v-if="replacementsUrl.length < 2000">
                <span>Or you can load all NAMES and VALUES by opening this url in any browser:</span> <br/>
                <pre class="my-url">{{replacementsUrl}}</pre>
              </p><br/>
              <p v-if="replacementNamesUrl.length < 2000">
                <span>Or you can load NAMES ONLY by opening this url in any browser:</span> <br/>
                <pre class="my-url">{{replacementNamesUrl}}</pre>
              </p>
            </div>
          </article>
        </div>
        <button class="modal-close is-large" aria-label="close" @click="saveModal.opened = false"></button>
      </div>

    </section>
  </div>
</template>

<script>
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import { saveAs } from "file-saver";

export default {
  name: 'Home',
  data () {
    return {
      files: [],
      replacements: [],
      saveModal : {
        opened: '',
      },
      error: '',
      index: -1,
    }
  },
  computed: {
    replacementsUrl () {
      var prefix = window.location.origin + '/load/'
      return prefix + encodeURIComponent(JSON.stringify(this.replacements))
    },
    replacementNamesUrl () {
      var prefix = window.location.origin + '/load/'
      var names = this.replacements.map(r => r.name)
      return prefix + encodeURIComponent(JSON.stringify(names))
    },
    config () {
      return this.$route.params.config
    },
    canGenerate () {
      return this.files.length
    },
    generating () {
      return this.index >=0
    },
  },
  watch: {
    config: function (val) {
      if (val) {
        this.loadConfig()
      }
    },
    index: function (val) {
      if (val >= 0) {
        this.generateOne()
      }
    },
  },
  methods: {
    addReplacement () {
      this.replacements.push({
        name: '',
        value: '',
      })
    },
    clearValues () {
      for(var replacement of this.replacements) {
        replacement.value = ''
      }
    },
    removeReplacements (index) {
      if (index == null) {
        this.replacements = []
      } else {
        this.replacements.splice(index, 1)
      }
    },
    saveReplacements () {
      if (this.config) {
        this.$router.push('/')
      }
      localStorage.setItem('replacements', JSON.stringify(this.replacements))
      this.saveModal.opened = true
    },
    loadConfig () {
      try {
        var replacements = JSON.parse(decodeURIComponent(this.config))
        if (!Array.isArray(replacements)) {
          throw 'not array'
        }
        if (!replacements.length) {
          return
        }
        if (typeof(replacements[0]) == 'string') {
          this.replacements = replacements.map(n => ({name: n, value: ''}))
        } else {
          this.replacements = replacements
        }
      } catch (error) {
        console.error(error)
        this.error = 'Failed to load replacements!'
      }
    },
    onFileChange (e) {
      var files = e.target.files || e.dataTransfer.files;
      if (!files.length)
        return;
      this.files = files;
    },
    generate () {
      if (!this.canGenerate || this.generating) {
        return
      }
      this.index = 0
    },
    generateOne () {
      console.log('generateOne', this.index)
      if (this.index >= this.files.length) {
        this.index = -1
        return
      }
      console.log('generateOne start')
      var replacements = {}
      for (const r of this.replacements) {
        replacements[r.name] = r.value
      }

      var file = this.files[this.index]

      var vm = this
      var reader = new FileReader();
      reader.onload = function(e) {
        // binary data
        console.log('binary', e.target.result);
        var content = e.target.result
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
        });
        doc.render(replacements);
        const out = doc.getZip().generate({
          type: "blob",
          mimeType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
        // Output the document using Data-URI
        saveAs(out, file.name);
        vm.index = vm.index + 1
      };
      reader.onerror = function(e) {
        // error occurred
        console.log('Error : ' + e.type);
        vm.error = 'Failed to generate some files'
        vm.index = -1
      };
      reader.readAsBinaryString(file);
    },
  },
  mounted () {
    if (this.config) {
      this.loadConfig()
    } else {
      var str = localStorage.getItem('replacements')
      if (str) {
        this.replacements = JSON.parse(str)
      }
    }
  },
}
</script>

<style scoped>
.my-url {
  word-break: break-all!important;
  white-space: normal;
}
</style>
