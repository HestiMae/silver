class wordUtil {

}

namespace wordUtil {
  export const bindTextPromise = (destinationName) => {
    return new Promise<Office.TextBinding>((resolve, reject) => {
      try {
        Office.context.document.bindings.addFromNamedItemAsync(destinationName, Office.BindingType.Text, { id: destinationName }, (result => {
          if (result.status == Office.AsyncResultStatus.Failed) {
            reject(result.error)
          } else {
            resolve(result.value)
          }
        }))
      } catch (error) {
        reject(error)
      }
    })
  }

  export const wordResultPromise = (promiseAll: (context: Word.RequestContext) => Promise<boolean[]>) => {
    return new Promise<boolean[]>(async (resolve, reject) => {
      await Word.run(async (context) => {
        resolve(promiseAll(context))
      })
    })
  }
}

export default wordUtil