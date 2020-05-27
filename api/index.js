module.exports = async (req, res) => {

  console.log('REQUEST body', req.body);
  console.log('REQUEST body', req.query);


  res.json({body: req});
}