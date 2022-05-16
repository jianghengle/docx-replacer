<template>
  <div>
    <nav class="navbar is-dark" role="navigation" aria-label="main navigation">
      <div class="container">
        <div class="navbar-brand">
          <router-link class="navbar-item is-size-4 has-text-weight-bold" :to="'/'">
            Docx Replacer
          </router-link>

          <a role="button" class="navbar-burger" aria-label="menu" aria-expanded="false"
            :class="{'is-active': menuActive}" @click="menuActive = !menuActive">
            <span aria-hidden="true"></span>
            <span aria-hidden="true"></span>
            <span aria-hidden="true"></span>
          </a>
        </div>

        <div id="navbarBasicExample" class="navbar-menu"  :class="{'is-active': menuActive}">
          <div class="navbar-start">
          </div>

          <div class="navbar-end">
          </div>
        </div>
      </div>
    </nav>
  </div>
</template>

<script>
import Vue from 'vue'

export default {
  name: 'Header',
  data () {
    return {
      menuActive: false
    }
  },
  computed: {
    
  },
  methods: {
    
  },
  mounted () {
  }
}
</script>